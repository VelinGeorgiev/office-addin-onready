import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import * as ReactDOM from "react-dom";

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */

Office.initialize = () => {
  console.log('initialize');

  Word.run(context => {

    console.log('Context initialized');
    console.log(context);

     // insert a paragraph at the end of the document.
     const paragraph = context.document.body.insertParagraph("initialize", Word.InsertLocation.end);

     // change the paragraph color to blue.
     paragraph.font.color = "red";

     context.sync();
  });
};

/* Render application after Office initializes */
Office.onReady(() => {
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);

  isOfficeInitialized = true;
  console.log('Office.onReady');

  Word.run(context => {

    console.log('Context Office.onReady triggered');
    console.log(context);

     // insert a paragraph at the end of the document.
     const paragraph = context.document.body.insertParagraph("Office.onReady", Word.InsertLocation.end);

     // change the paragraph color to blue.
     paragraph.font.color = "green";

     context.sync();
  });
  render(App);
});

/* Initial render showing a progress bar */
render(App);

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
