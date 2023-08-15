import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import Home from './components/Home';
import { IHomeProps } from './components/IHome';

export interface IHomeWebPart {
  context: WebPartContext;
}

export default class HomeWebPart extends BaseClientSideWebPart<IHomeWebPart> {

  public render(): void {
    const element: React.ReactElement<IHomeProps> = React.createElement(
      Home,
      {
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
