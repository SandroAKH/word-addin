import App from "./components/App"
import { FluentProvider, webLightTheme } from "@fluentui/react-components"
import * as React from "react"
import * as ReactDOM from "react-dom"

/* global document, Office, module, require */

let isOfficeInitialized = false

const title = "Contoso Task Pane Add-in"

const render = Component => {
  ReactDOM.render(
    <FluentProvider theme={webLightTheme}>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </FluentProvider>,
    document.getElementById("container")
  )
}

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true
  render(App)
})

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default
    render(NextApp)
  })
}
