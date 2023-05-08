import * as React from 'react';
import styles from './Taskscount.module.scss';
import { ITaskscountProps } from './ITaskscountProps';

import {
  SPFI
} from '@pnp/sp';
import { getSP } from '../../../pnpjsConfig';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import "@pnp/sp/site-users/web";

export interface IAuthor {
  Id: number;
  Title: string;
}

export interface IUrl {
  Url: string;
}

export interface IFile {
  Id?: number;
  Title?: string;
  TaskId?: string;
  ProductName?: string;
  Category?: string;
  SubCategory?: string;
  Author?: IAuthor;
  Status?: string;
  Created?: Date;
  Modified?: Date;
  URL?: IUrl;
  AssignedToId?: number[];
  DelegateUserId?: number[];
}

export interface IPnPjsV3State {
  items: IFile[];
}

export class PnPjsV3State implements IPnPjsV3State {
  constructor(
    public items: IFile[] = []
  ) { }
}

export default class Taskscount extends React.Component<ITaskscountProps, IPnPjsV3State, {}> {

  private LIST_NAME: string = "Tasks";
  private _sp: SPFI;

  constructor(props: ITaskscountProps) {
    super(props);
    this.state = new PnPjsV3State();
    this._sp = getSP(props.context);
  }

  public componentDidMount(): void {
    this._readAllItems().then((items) => { console.log() }).catch(err => console.error(err));
    this._readAllItems2().then((items) => { console.log() }).catch(err => console.error(err));
    this._readAllItems3().then((items) => { console.log() }).catch(err => console.error(err));
    this._readAllItems4().then((items) => { console.log() }).catch(err => console.error(err));
  }

  private _readAllItems = async (): Promise<void> => {
    let items: any[] =[];
    try {

       const user = await this._sp.web.currentUser();

       const userId = await user.Id;

      const response: any[] = await this._sp.web.lists.getByTitle(this.LIST_NAME).items
      .orderBy("ID", false)
      .filter("Status eq 'Pending' and (AssignedToId ne null or DelegateUserId ne null)")
      .select("AssignedToId", "DelegateUserId", "Status")
      .top(100)();

      items = await (response.filter((item: any) => {

        return (item.Status == 'Pending' && item.AssignedToId != null && item.AssignedToId.includes(userId )
        || item.Status == 'Pending' && item.DelegateUserId != null && item.DelegateUserId.includes(userId))

      }));

      this.setState({ items });

    } catch (err) {
      console.error(`Error - ${JSON.stringify(err)} - `);
    }
  }

  private _readAllItems2 = async (): Promise<void> => {
    let items: any[] =[];
    try {

       const user = await this._sp.web.currentUser();

       const userId = await user.Id;

      const response2: any[] = await this._sp.web.lists.getByTitle(this.LIST_NAME).items
      .orderBy("ID", false)
      .filter("Status eq 'Pending' and (AssignedToId ne null or DelegateUserId ne null)")
      .select("AssignedToId", "DelegateUserId", "Status")      
      .top(500)()

      items = await (response2.filter((item: any) => {

        return (item.Status == 'Pending' && item.AssignedToId != null && item.AssignedToId.includes(userId )
        || item.Status == 'Pending' && item.DelegateUserId != null && item.DelegateUserId.includes(userId))

      }));

      this.setState({ items });

    } catch (err) {
      console.error(`Error - ${JSON.stringify(err)} - `);
    }
  }

  private _readAllItems3 = async (): Promise<void> => {
    let items: any[] =[];
    try {

       const user = await this._sp.web.currentUser();

       const userId = await user.Id;
      const response3: any[] = await this._sp.web.lists.getByTitle(this.LIST_NAME).items
      .orderBy("ID", false)
      .filter("Status eq 'Pending' and (AssignedToId ne null or DelegateUserId ne null)")
      .select("AssignedToId", "DelegateUserId", "Status")      
      .top(5000)()

      items = await (response3.filter((item: any) => {

        return (item.Status == 'Pending' && item.AssignedToId != null && item.AssignedToId.includes(userId )
        || item.Status == 'Pending' && item.DelegateUserId != null && item.DelegateUserId.includes(userId))

      }));

      this.setState({ items });

    } catch (err) {
      console.error(`Error - ${JSON.stringify(err)} - `);
    }
  }

  private _readAllItems4 = async (): Promise<void> => {
    let items: any[] =[];
    try {

       const user = await this._sp.web.currentUser();

       const userId = await user.Id;

      const response3: any[] = await this._sp.web.lists.getByTitle(this.LIST_NAME).items
      .select("AssignedToId", "DelegateUserId", "Status")
      .getAll()     

      items = await (response3.filter((item: any) => {

        return (item.Status == 'Pending' && (item.AssignedToId != null && item.AssignedToId.includes(userId ))
        || item.Status == 'Pending' && (item.DelegateUserId != null && item.DelegateUserId.includes(userId)))

      }));

      this.setState({ items });

    } catch (err) {
      console.error(`Error - ${JSON.stringify(err)} - `);
    }
  }

  public render(): React.ReactElement<ITaskscountProps> {

    return (
        
      <div className={styles.container} >
        <a style={{textDecoration: 'none'}} className={styles.links} href="https://banglalinkdigitalcomm.sharepoint.com/sites/vloungeonline/SitePages/Pending-Task.aspx">
          <div className={styles.flexContainer} >
          
            <div className={styles.flexChild} >
              <div className={styles.welcomeImage}></div>
            </div>

            <div className={styles.flexChildTxt}>
              <div >
                <span className={styles.text} ><b>Pending Task(s): </b></span>
              </div>
            </div>

            <div className={styles.flexCount}>                  
              <div className={styles.fieldsetDiv}>
                <fieldset className={styles.fieldsetBrdr}>
                  <span className={styles.text}><b>{this.state.items.length} </b></span>
                </fieldset>                                          
              </div> 
            </div>

          </div>
          
        </a>            
      </div>
        
    );
  }
  
}
