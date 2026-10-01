import * as React from 'react'
import { ExternalIcon } from '../icons'
import { BANNER_LINKS, STUDIO_URL } from '../../constants/links'

type BannerMenuItem = {
  key: string
  icon: string
  label: string
  description: string
  href: string
  external?: boolean
}

const BANNER_MENU_ITEMS: BannerMenuItem[] = [
  { key: 'work-request', icon: '📋', label: '업무 요청', description: 'Agent 관련 신청 및 요청', href: BANNER_LINKS.workRequest, external: true },
  { key: 'operation-status', icon: '📊', label: '운영 현황', description: '공동 운영 과제 및 문의 현황 확인', href: BANNER_LINKS.operationStatus, external: true },
  { key: 'operation-guide', icon: '📚', label: '운영 정보', description: '과제 등록 · 인수인계 · 운영 가이드', href: BANNER_LINKS.operationGuide, external: true },
  { key: 'studio', icon: '✏️', label: '게시물 작성 Studio', description: '섹션을 조합해 운영 게시물 작성', href: STUDIO_URL },
]

function BannerMenuCard({ item }: { item: BannerMenuItem }) {
  return (
    <a
      className="banner-menu-item"
      href={item.href}
      target={item.external ? '_blank' : undefined}
      rel={item.external ? 'noreferrer' : undefined}
    >
      <span className="banner-menu-item-icon">{item.icon}</span>
      <div className="banner-menu-item-copy">
        <div className="banner-menu-item-title clamp-1" title={item.label}>{item.label}</div>
        <div className="banner-menu-item-description clamp-1" title={item.description}>{item.description}</div>
      </div>
      <span className="banner-menu-item-external"><ExternalIcon /></span>
    </a>
  )
}

export default function BannerMenu() {
  return (
    <div className="banner-menu-list">
      {BANNER_MENU_ITEMS.map(item => <BannerMenuCard key={item.key} item={item} />)}
    </div>
  )
}
