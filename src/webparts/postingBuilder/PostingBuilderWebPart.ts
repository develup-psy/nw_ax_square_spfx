import * as React from 'react'
import * as ReactDom from 'react-dom'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'

import PostingBuilderApp from './components/PostingBuilder'
import { initSP } from './services/pnp'
import { getCurrentUser } from './services/user'
import { createSharePointPostingBuilderRepository } from './services/postingBuilderSharePoint'

import '../../shared/styles/fonts.css'

export interface IPostingBuilderWebPartProps {}

export default class PostingBuilderWebPart
  extends BaseClientSideWebPart<IPostingBuilderWebPartProps> {

  protected async onInit(): Promise<void> {
    await super.onInit()

    initSP(this.context)
  }

  public render(): void {
    const user =
      getCurrentUser(this.context)

    const repository =
      createSharePointPostingBuilderRepository()

    const squareUrl =
      `${this.context.pageContext.web.serverRelativeUrl}` +
      `/SitePages/NW_AX_Square.aspx?env=WebView`

    const element =
      React.createElement(
        PostingBuilderApp,
        {
          user,
          repository,

          onBackToSquare: () => {
            window.location.href = squareUrl
          },
        }
      )

    ReactDom.render(
      element,
      this.domElement
    )
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(
      this.domElement
    )
  }
}