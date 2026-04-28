public render(): void {
  const element: React.ReactElement<any> = React.createElement(
    AuthorizedRequestor as any,
    {
      context: this.context,
      displayMode: this.displayMode,
      onSave: this._onSave,
      onClose: this._onClose
    }
  );

  ReactDOM.render(element, this.domElement);
}