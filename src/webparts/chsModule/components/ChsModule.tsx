import * as React from 'react';
import styles from './ChsModule.module.scss';
import { IChsModuleProps } from './IChsModuleProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IList, Web } from "@pnp/sp/presets/all";
import { BrowserRouter as Router, Switch, Route, Link, HashRouter, match, useParams, Redirect } from 'react-router-dom';
import CHSCreation from './CHSCreation/CHSCreation'
import { sp } from '@pnp/sp';
import '@pnp/sp/lists';
import '@pnp/sp/items';
require('../assets/style.css');
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
export default class ChsModule extends React.Component < IChsModuleProps, {} > {
  public render(): React.ReactElement<IChsModuleProps> {
    
    return(
      <section >
        <div id='divLoader' className={'divLoader'}></div>
        <HashRouter>
          <div className='wrapper'>
            {}
            <div className='main'>
              <div className='content'>
                <Switch>
                  <Route path="/" render={() => <CHSCreation  {...this.props} />} />
                </Switch>
              </div>
            </div>
          </div>
        </HashRouter>
      </section>
    );
  }
}