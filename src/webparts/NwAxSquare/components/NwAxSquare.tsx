import * as React from 'react'
import Header from './components/Header'
import Main from './components/main/Main'
import HotNoticePopup from './components/hot-notice/HotNoticePopup'
import Footer from './components/Footer'
import NwAxSquareTutorial from './components/NwAxSquareTutorial'
import './styles.css'

export default function App() {
  return (
    <div className="app-root">
      <Header />
      <Main />
      <Footer />

      <HotNoticePopup />
      <NwAxSquareTutorial />
    </div>
  )
}