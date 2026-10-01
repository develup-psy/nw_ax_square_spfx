import * as React from 'react'
import { useAxSquare } from '../../context/AxSquareContext'

interface BannerMenuItem {
  key: string
  icon: string
  label: string
  description: string
  href: string
  external?: boolean
}

/**
 * SharePoint에서 데이터를 받지 못했을 때 사용하는 가데이터
 */
const FALLBACK_BANNER_MENU: BannerMenuItem[] = [
  {
    key: 'work-request',
    icon: '📋',
    label: '업무 요청',
    description: 'Agent 관련 신청 및 요청',
    href:
      'https://lgucorp.atlassian.net/jira/core/projects/NWAX/board?filter=&groupBy=none',
    external: true,
  },
  {
    key: 'operation-status',
    icon: '📊',
    label: '운영 현황',
    description: '공동 운영 과제 및 문의 현황 확인',
    href:
      'https://lgucorp.atlassian.net/wiki/spaces/yuE2EYBu1YuB/pages/1487077824',
    external: true,
  },
  {
    key: 'operation-guide',
    icon: '📚',
    label: '운영 정보',
    description: '과제 등록 · 인수인계 · 운영 가이드',
    href:
      'https://lgucorp.atlassian.net/wiki/spaces/yuE2EYBu1YuB/folder/1489014139/03.',
    external: true,
  },
  {
    key: 'studio',
    icon: '✏️',
    label: '게시물 작성 Studio',
    description: '섹션을 조합해 운영 게시물 작성',
    href:
      'https://lgupluscorp.sharepoint.com/sites/NW_AX_SQUARE/SitePages/NW-AX-Square-Studio.aspx',
    external: false,
  },
]

const COPILOT_URL =
  'https://m365.cloud.microsoft/chat/agent/T_e3042231-6b5a-80b1-d0a4-e3914c7c0a90.ab993c79-d288-45e7-aaf1-51c66e081f69.gpt.4b89dee4-96b9-4086-83ac-d60040d8b77f'

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

function BotIcon(): React.ReactElement {
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
      <rect x="3" y="11" width="18" height="10" rx="2" />
      <circle cx="12" cy="5" r="2" />
      <path d="M12 7v4" />
    </svg>
  )
}

function CopilotIcon(): React.ReactElement {
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
      <path d="M12 3v18" />
      <path d="M3 12h18" />
      <path d="m5.6 5.6 12.8 12.8" />
      <path d="m18.4 5.6-12.8 12.8" />
    </svg>
  )
}

function BannerMenuItemView({
  item,
}: {
  item: BannerMenuItem
}): React.ReactElement {
  return (
    <a
      className="banner-menu-item"
      href={item.href}
      target={item.external ? '_blank' : undefined}
      rel={item.external ? 'noreferrer' : undefined}
    >
      <span className="banner-menu-item-icon">
        {item.icon}
      </span>

      <div className="banner-menu-item-copy">
        <div
          className="banner-menu-item-title clamp-1"
          title={item.label}
        >
          {item.label}
        </div>

        <div
          className="banner-menu-item-description clamp-1"
          title={item.description}
        >
          {item.description}
        </div>
      </div>

      <span className="banner-menu-item-external">
        <ExternalLinkIcon />
      </span>
    </a>
  )
}

function BannerMenu(): React.ReactElement {
  const { getPostsBySection } = useAxSquare()

  const managedPosts =
    getPostsBySection('BANNER_MENU')

  /**
   * 기존 메뉴 데이터를 기준으로 UI는 그대로 유지합니다.
   *
   * SharePoint에 동일한 Title이 있고 href가 존재하면
   * SharePoint URL을 사용합니다.
   *
   * 데이터가 없거나 href가 비어 있으면
   * FALLBACK_BANNER_MENU의 기존 URL을 사용합니다.
   */
  const menuItems = FALLBACK_BANNER_MENU.map(
    (fallbackItem) => {
      const managedPost = managedPosts.find(
        (post) =>
          post.title.trim() ===
          fallbackItem.label.trim()
      )

      return {
        ...fallbackItem,
        href:
          managedPost?.href?.trim() ||
          fallbackItem.href,
      }
    }
  )

  return (
    <div className="banner-menu-list">
      {menuItems.map((item) => (
        <BannerMenuItemView
          key={item.key}
          item={item}
        />
      ))}
    </div>
  )
}

export default function Banner(): React.ReactElement {
  return (
    <section className="banner-section">
      <div className="page-container banner-grid">
        <div className="banner-content">
          <div className="banner-team-label">
            <span>LG U+</span>

            <span className="banner-team-divider" />

            <span className="clamp-1">
              NW AX추진팀
            </span>
          </div>

          <div className="banner-title-row">
            <h1
              className="banner-title clamp-2"
              title="NW AX 운영스퀘어"
            >
              NW AX 운영스퀘어
            </h1>

            <a
              className="banner-copilot-link banner-copilot-mobile"
              href={COPILOT_URL}
              target="_blank"
              rel="noreferrer"
              title="운영센터 코파일럿 챗봇"
            >
              <CopilotIcon />
              <span>챗봇</span>
            </a>
          </div>

          <div className="banner-subtitle clamp-1">
            NW AX추진팀 · Agent 운영을 위한 통합 관리 플랫폼
          </div>

          <p className="banner-description clamp-3">
            운영 업무에 필요한 다양한 콘텐츠와 정보를 제공하여
            <br />
            보다 안정적이고 체계적인 업무 수행을 지원하고자 운영
            스퀘어를 오픈했습니다.
          </p>

          <a
            className="banner-copilot-link banner-copilot-desktop"
            href={COPILOT_URL}
            target="_blank"
            rel="noreferrer"
          >
            <CopilotIcon />
            운영스퀘어 챗봇
          </a>
        </div>

        <div className="banner-menu-card">
          <div className="banner-menu-shell">
            <div className="banner-menu-header">
              <span className="banner-menu-header-icon">
                <BotIcon />
              </span>

              <span className="clamp-1">
                NW AX 운영스퀘어
              </span>
            </div>

            <BannerMenu />
          </div>

          <div className="banner-team-badge">
            NW AX추진팀
          </div>
        </div>
      </div>
    </section>
  )
}