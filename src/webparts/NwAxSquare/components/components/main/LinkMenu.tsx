import * as React from 'react'
import { useAxSquare } from '../../context/AxSquareContext'

interface LinkMenuItem {
  key: string
  icon: string
  title: string
  desc: string
  href: string
}

/**
 * 업무 요청 - SharePoint 데이터를 받지 못할 경우 사용할 가데이터
 */
const FALLBACK_REQUEST_ITEMS: LinkMenuItem[] = [
  {
    key: 'new-environment',
    icon: '📄',
    title: '신규 개발기/운영기 신청',
    desc: '신규 Agent 개발 및 운영 환경 신청',
    href:
      'https://lgucorp.atlassian.net/jira/core/projects/NWAX/form/3867',
  },
  {
    key: 'schedule',
    icon: '📅',
    title: '과제 스케줄 등록/변경 신청',
    desc: 'Agent 수행 스케줄 등록 및 변경 요청',
    href:
      'https://lgucorp.atlassian.net/jira/core/projects/NWAX/form/3866',
  },
  {
    key: 'shared-operation',
    icon: '🔄',
    title: '공동 운영 전환 신청',
    desc: '개별 운영 과제를 공동 운영 체계로 전환 요청',
    href:
      'https://lgucorp.atlassian.net/jira/core/projects/NWAX/form/3860',
  },
  {
    key: 'faq',
    icon: '💬',
    title: '문의사항 등록 FAQ',
    desc: '운영 관련 문의 및 FAQ 접수',
    href:
      'https://lgucorp.atlassian.net/jira/core/projects/NWAX/form/3865',
  },
]

/**
 * 운영 현황 - SharePoint 데이터를 받지 못할 경우 사용할 가데이터
 */
const FALLBACK_STATUS_ITEMS: LinkMenuItem[] = [
  {
    key: 'shared-status',
    icon: '📊',
    title: '공동 운영과제 현황',
    desc: '공동 운영 중인 Agent 목록 및 현황 확인',
    href:
      'https://lgucorp.atlassian.net/wiki/spaces/yuE2EYBu1YuB/pages/1487077824',
  },
  {
    key: 'inquiry-status',
    icon: '📋',
    title: '문의사항 현황',
    desc: '전체 문의 목록 및 처리 상태 확인',
    href:
      'https://lgucorp.atlassian.net/wiki/spaces/yuE2EYBu1YuB/pages/1486589820',
  },
]

function ExternalLinkIcon(): React.ReactElement {
  return (
    <svg
      width="16"
      height="16"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      <path d="M15 3h6v6" />
      <path d="M10 14 21 3" />
      <path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" />
    </svg>
  )
}

function LinkCard({
  item,
}: {
  item: LinkMenuItem
}): React.ReactElement {
  return (
    <a
      className="link-card"
      href={item.href}
      target="_blank"
      rel="noreferrer"
    >
      <span className="link-card-icon">
        {item.icon}
      </span>

      <div className="link-card-copy">
        <div
          className="link-card-title clamp-2"
          title={item.title}
        >
          {item.title}
        </div>

        <div
          className="link-card-description clamp-2"
          title={item.desc}
        >
          {item.desc}
        </div>
      </div>

      <div className="link-card-action">
        바로가기
        <ExternalLinkIcon />
      </div>
    </a>
  )
}

function LinkSectionLabel({
  children,
}: {
  children: React.ReactNode
}): React.ReactElement {
  return (
    <div className="link-section-label">
      <span className="link-section-label-bar" />

      <span className="clamp-1">
        {children}
      </span>
    </div>
  )
}

/**
 * SharePoint URL과 fallback URL을 병합합니다.
 *
 * UI의 제목 / 설명 / 아이콘은 기존 소스를 유지하고
 * URL만 SharePoint에서 가져옵니다.
 */
function applyManagedUrls(
  fallbackItems: LinkMenuItem[],
  managedPosts: ReturnType<
    ReturnType<typeof useAxSquare>['getPostsBySection']
  >
): LinkMenuItem[] {
  return fallbackItems.map((fallbackItem) => {
    const managedPost = managedPosts.find(
      (post) =>
        post.title.trim() ===
        fallbackItem.title.trim()
    )

    return {
      ...fallbackItem,
      href:
        managedPost?.href?.trim() ||
        fallbackItem.href,
    }
  })
}

export default function LinkMenu(): React.ReactElement {
  const { getPostsBySection } = useAxSquare()

  const managedPosts =
    getPostsBySection('LINK_MENU')

  const requestItems = applyManagedUrls(
    FALLBACK_REQUEST_ITEMS,
    managedPosts
  )

  const statusItems = applyManagedUrls(
    FALLBACK_STATUS_ITEMS,
    managedPosts
  )

  return (
    <section
      id="link-menu"
      className="link-menu-section"
    >
      <div className="page-container">
        <div className="section-heading link-menu-heading">
          <h2 className="section-title clamp-2">
            운영 업무 메뉴
          </h2>

          <p className="section-description clamp-2">
            아래 항목을 클릭하면 해당 페이지로 바로 이동합니다.
          </p>
        </div>

        <div className="link-menu-columns">
          <div className="link-menu-group">
            <LinkSectionLabel>
              업무 요청
            </LinkSectionLabel>

            <div className="link-card-grid">
              {requestItems.map((item) => (
                <LinkCard
                  key={item.key}
                  item={item}
                />
              ))}
            </div>
          </div>

          <div className="link-menu-group">
            <LinkSectionLabel>
              운영 현황
            </LinkSectionLabel>

            <div className="link-card-grid">
              {statusItems.map((item) => (
                <LinkCard
                  key={item.key}
                  item={item}
                />
              ))}
            </div>
          </div>
        </div>
      </div>
    </section>
  )
}