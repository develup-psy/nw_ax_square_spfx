import { WebPartContext } from '@microsoft/sp-webpart-base'
import type { StudioUser } from '../components/studio/types'

/**
 * 별도 Graph 호출 없이 SPFx PageContext에서 현재 Microsoft 365 사용자 정보를 가져옵니다.
 * userId는 가능하면 Entra ID Object ID를 사용하고, 없는 환경에서는 email/loginName으로 fallback합니다.
 */
export function getCurrentUser(context: WebPartContext): StudioUser {
  const pageUser = context.pageContext.user
  const aadUserId = context.pageContext.aadInfo?.userId?.toString()
  const fallbackId = pageUser.email || pageUser.loginName

  if (!aadUserId && !fallbackId) {
    throw new Error('현재 Microsoft 사용자 ID를 확인할 수 없습니다.')
  }

  return {
    id: aadUserId || fallbackId,
    name: pageUser.displayName || '사용자',
    email: pageUser.email || '',
  }
}
