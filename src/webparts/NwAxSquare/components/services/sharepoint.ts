import { getSP } from './pnp'
import '@pnp/sp/attachments'

export type AxSquarePost = {
  id: number
  title: string
  section: 'INFORMATION' | 'NOTICE' | 'INSIGHT' | 'HOT_NOTICE' | string
  tag: string
  summary: string
  href: string
  publishedAt: string
  sortOrder: number
  isFeatured: boolean
  isVisible: boolean
  contentImage: string
}

type HyperlinkValue =
  | string
  | {
      Url?: string
      Description?: string
    }

type SharePointPostItem = {
  Id: number
  Title?: string
  section?: string
  tag?: string
  summary?: string
  href?: HyperlinkValue
  published_at?: string
  sort_order?: number
  is_featured?: boolean
  is_visible?: boolean
  content_image?: unknown
}

type SharePointAttachmentInfo = {
  FileName?: string
  ServerRelativeUrl?: string
}

type JsonObject = Record<string, unknown>

function asString(value: unknown): string {
  if (
    value === undefined ||
    value === null
  ) {
    return ''
  }

  return String(value).trim()
}

function parseObject(
  value: unknown
): JsonObject | null {
  if (!value) return null

  if (typeof value === 'object') {
    return value as JsonObject
  }

  const raw = asString(value)

  if (!raw) return null

  try {
    const parsed = JSON.parse(raw)

    return (
      parsed &&
      typeof parsed === 'object'
    )
      ? parsed as JsonObject
      : null
  } catch {
    return null
  }
}

/**
 * 일반 Hyperlink 컬럼 및
 * 기존 SharePoint Image URL 형식 처리
 */
function normalizeUrl(
  value: unknown
): string {
  if (!value) return ''

  const objectValue =
    parseObject(value)

  if (objectValue) {
    /**
     * 기존 Image 컬럼 형식
     */
    if (
      objectValue.serverUrl &&
      objectValue.serverRelativeUrl
    ) {
      return (
        `${asString(objectValue.serverUrl)}` +
        `${asString(
          objectValue.serverRelativeUrl
        )}`
      )
    }

    /**
     * Hyperlink / URL 형식
     */
    return asString(
      objectValue.Url ??
      objectValue.url ??
      objectValue.serverRelativeUrl ??
      objectValue.serverUrl
    )
  }

  /**
   * 일반 문자열 URL
   */
  return asString(value)
}

function toAbsoluteUrl(
  value: string
): string {
  const url = value.trim()

  if (!url) return ''

  if (/^https?:\/\//i.test(url)) {
    return url
  }

  if (url.startsWith('/')) {
    return (
      `${window.location.origin}${url}`
    )
  }

  return url
}

/**
 * content_image 전용 처리
 *
 * 1.
 * 기존 형식:
 *
 * {
 *   serverUrl,
 *   serverRelativeUrl
 * }
 *
 * 2.
 * 신규 형식:
 *
 * {
 *   fileName:
 *   "Reserved_ImageAttachment_....png"
 * }
 *
 * 신규 형식은 실제 이미지가
 * List Item attachment에 있기 때문에
 * attachmentFiles에서 실제 URL 조회
 */
async function resolveContentImage(
  itemId: number,
  value: unknown
): Promise<string> {
  if (!value) return ''

  const objectValue =
    parseObject(value)

  /**
   * 일반 URL 문자열
   */
  if (!objectValue) {
    return toAbsoluteUrl(
      normalizeUrl(value)
    )
  }

  /**
   * 기존 SharePoint Image 컬럼
   *
   * serverRelativeUrl 등이 있으면
   * 추가 API 호출 없이 그대로 사용
   */
  const directUrl =
    normalizeUrl(objectValue)

  if (directUrl) {
    return toAbsoluteUrl(directUrl)
  }

  /**
   * 신규 Reserved_ImageAttachment 형식
   */
  const fileName =
    asString(objectValue.fileName)

  if (!fileName) {
    return ''
  }

  try {
    const sp = getSP()

    const attachments =
      await sp.web.lists
        .getByTitle(
          'ax_square_posts'
        )
        .items
        .getById(itemId)
        .attachmentFiles
        .select(
          'FileName',
          'ServerRelativeUrl'
        )() as SharePointAttachmentInfo[]

    const matched =
      attachments.find(
        attachment =>
          attachment.FileName ===
          fileName
      )

    if (
      !matched?.ServerRelativeUrl
    ) {
      console.warn(
        '[NW AX Square] content_image attachment를 찾지 못했습니다.',
        {
          itemId,
          fileName,
        }
      )

      return ''
    }

    return toAbsoluteUrl(
      matched.ServerRelativeUrl
    )
  } catch (error) {
    console.error(
      '[NW AX Square] content_image 조회 실패',
      {
        itemId,
        fileName,
        error,
      }
    )

    return ''
  }
}

export async function fetchAxSquarePosts():
Promise<AxSquarePost[]> {
  const sp = getSP()

  const items = await sp.web.lists
  .getByTitle('ax_square_posts')
  .items
  .select(
    'Id',
    'Title',
    'section',
    'tag',
    'summary',
    'href',
    'published_at',
    'sort_order',
    'is_featured',
    'is_visible',
    'content_image'
  )
  .filter('is_visible eq 1')
  .top(5000)() as SharePointPostItem[];

  return Promise.all(
    items.map(
      async item => ({
        id:
          item.Id,

        title:
          item.Title || '',

        section:
          (
            item.section || ''
          ).toUpperCase(),

        tag:
          item.tag || '',

        summary:
          item.summary || '',

        href:
          normalizeUrl(
            item.href
          ),

        publishedAt:
          item.published_at || '',

        sortOrder:
          item.sort_order ??
          9999,

        isFeatured:
          item.is_featured ??
          false,

        isVisible:
          item.is_visible ??
          true,

        contentImage:
          await resolveContentImage(
            item.Id,
            item.content_image
          ),
      })
    )
  )
}