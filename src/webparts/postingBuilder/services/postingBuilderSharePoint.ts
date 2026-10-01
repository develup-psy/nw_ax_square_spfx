import { getSP } from './pnp'
import type { BuilderProject } from '../components/types/project'
import type {
  CreateStudioPostInput,
  PostingBuilderRepository,
  StudioPost,
  StudioPostSummary,
} from '../components/studio/types'

const LIST_NAME = 'ax_square_builder_posts'

const CONTENT_SAFE_LIMIT = 60000

function serializeProject(content: BuilderProject): string {
  const json = JSON.stringify(content)
  if (json.length > CONTENT_SAFE_LIMIT) {
    throw new Error(
      '게시물 데이터가 SharePoint content 컬럼의 안전한 저장 크기를 초과했습니다. ' +
      '업로드 이미지가 JSON에 Base64로 포함되어 있는지 확인해주세요.'
    )
  }
  return json
}

type SharePointBuilderPostItem = {
  Id: number
  Title?: string
  post_id?: string
  user_id?: string
  content?: string
  Modified?: string
}

function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''")
}

function parseProject(content: string | undefined): BuilderProject {
  if (!content) throw new Error('게시물 content가 비어 있습니다.')
  const parsed = JSON.parse(content) as BuilderProject
  if (!parsed || !Array.isArray(parsed.sections)) {
    throw new Error('저장된 게시물 JSON 형식이 올바르지 않습니다.')
  }
  return parsed
}

export function createSharePointPostingBuilderRepository(): PostingBuilderRepository {
  return {
    async listPosts(userId: string): Promise<StudioPostSummary[]> {
      const sp = getSP()
      const safeUserId = escapeODataValue(userId)

      const items = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .select('Id', 'Title', 'post_id', 'Modified')
        .filter(`user_id eq '${safeUserId}'`)
        .orderBy('Modified', false)
        .top(200)() as SharePointBuilderPostItem[]

      return items.map(item => ({
        itemId: item.Id,
        postId: item.post_id || '',
        title: item.Title || '제목 없는 게시물',
        modified: item.Modified || '',
      }))
    },

    async getPost(userId: string, postId: string): Promise<StudioPost | null> {
      const sp = getSP()
      const safePostId = escapeODataValue(postId)

      // post_id는 Unique + Indexed로 두므로 단일 조건으로 먼저 찾고,
      // 조회 후 user_id를 다시 확인합니다.
      const items = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .select('Id', 'Title', 'post_id', 'user_id', 'content', 'Modified')
        .filter(`post_id eq '${safePostId}'`)
        .top(1)() as SharePointBuilderPostItem[]

      if (!items.length) return null
      const item = items[0]
      if ((item.user_id || '') !== userId) return null

      return {
        itemId: item.Id,
        postId: item.post_id || '',
        userId: item.user_id || '',
        title: item.Title || '제목 없는 게시물',
        content: parseProject(item.content),
        modified: item.Modified || '',
      }
    },

    async createPost(input: CreateStudioPostInput): Promise<StudioPost> {
      const sp = getSP()
      const result = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .add({
          Title: input.title,
          post_id: input.postId,
          user_id: input.userId,
          content: serializeProject(input.content),
        })

      return {
        itemId: result.Id,
        postId: input.postId,
        userId: input.userId,
        title: input.title,
        content: input.content,
        modified: new Date().toISOString(),
      }
    },

    async updatePost(itemId: number, title: string, content: BuilderProject): Promise<void> {
      const sp = getSP()
      await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .getById(itemId)
        .update({
          Title: title || '제목 없는 게시물',
          content: serializeProject(content),
        })
    },


    async deletePost(userId: string, postId: string): Promise<void> {
      const sp = getSP()
      const safePostId = escapeODataValue(postId)

      // post_id는 Unique + Indexed이므로 먼저 해당 Item을 찾고,
      // 실제 삭제 전에 user_id를 다시 비교하여 다른 사용자의 Item 삭제를 막습니다.
      const items = await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .select('Id', 'post_id', 'user_id')
        .filter(`post_id eq '${safePostId}'`)
        .top(1)() as SharePointBuilderPostItem[]

      if (!items.length) {
        throw new Error('삭제할 게시물을 찾지 못했습니다.')
      }

      const item = items[0]
      if ((item.user_id || '') !== userId) {
        throw new Error('현재 사용자가 삭제할 수 있는 게시물이 아닙니다.')
      }

      // delete()가 아니라 recycle()을 사용하여 SharePoint 휴지통으로 이동합니다.
      await sp.web.lists
        .getByTitle(LIST_NAME)
        .items
        .getById(item.Id)
        .recycle()
    },
  }
}
