import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import DocumentAnalyzer from './components/DocumentAnalyzer';

export default class DocumentAnalyzerWebPart extends BaseClientSideWebPart<{}> {

  public render(): void {
    const element = React.createElement(DocumentAnalyzer, {
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }
}