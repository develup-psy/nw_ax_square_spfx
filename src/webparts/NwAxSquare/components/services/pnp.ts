import { spfi, SPFI, SPFx } from '@pnp/sp'
import { WebPartContext } from '@microsoft/sp-webpart-base'

import '@pnp/sp/webs'
import '@pnp/sp/lists'
import '@pnp/sp/items'

let _sp: SPFI | null = null

export function initSP(context: WebPartContext): void {
  _sp = spfi().using(SPFx(context))
}

export function getSP(): SPFI {
  if (!_sp) {
    throw new Error('PnPjs가 초기화되지 않았습니다. initSP()를 먼저 호출해주세요.')
  }
  return _sp
}
