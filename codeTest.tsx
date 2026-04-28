const element: React.ReactElement<IAuthorizedRequestorProps> =
  React.createElement(AuthorizedRequestor, {
    context: this.context,
    displayMode: this.displayMode,
    onSave: this._onSave,
    onClose: this._onClose
  });

ReactDOM.render(element, this.domElement);

import AuthorizedRequestor from "./components/AuthorizedRequestor";
import { IAuthorizedRequestorProps } from "./components/IAuthorizedRequestorProps";