//loader
import { SPComponentLoader } from '@microsoft/sp-loader';

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ApagaSliderWebPart.module.scss';
import * as strings from 'ApagaSliderWebPartStrings';

//pnp
import { spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"

//bootstrap
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');

export interface IApagaSliderWebPartProps {
  description: string;
}

export default class ApagaSliderWebPart extends BaseClientSideWebPart<IApagaSliderWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.apagaSlider}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">              
              <div class="row">
              <div class="col-md-8">
                <h2>Apagar slider</h2>
                <div class="form-group">
                  <input type="text" id="txtTitulo" placeholder="Título do slider" class="form-control"/>
                  <input type="text" id="txtLink" placeholder="https://www.google.com.br" class="form-control"/>
                  <br/>
                  <button type="button" class="btn btn-danger apagarSlider">Excluir</button>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>`;

    this.setButtonsEventHandlers();
    this.pageLoad();
  }

  private pageLoad() : void
  {
      var origem = this.getQueryStringParameter("idSlider");      
      if(origem != "")
      {
        var idItem = parseInt(origem);
        sp.web.lists.getByTitle(this.properties.description).items.getById(idItem).get().then((item: any) => {
          (<HTMLInputElement>document.getElementById('txtTitulo')).value = item.Title;
          (<HTMLInputElement>document.getElementById('txtLink')).value = item.Link;
        });
      } 
  }

  private setButtonsEventHandlers(): void {
    const webPart: ApagaSliderWebPart = this;
    this.domElement.querySelector('button.apagarSlider').addEventListener('click', () => { webPart.apagarSlider(); });
  }

  private apagarSlider(): void {
    var origem = this.getQueryStringParameter("idSlider");
    var idSlider = parseInt(origem);

    sp.web.lists.getByTitle(this.properties.description).items.getById(idSlider).delete();    
    console.log("Slider ID : "+ idSlider + " Apagado !");
    window.location.href = this.context.pageContext.web.absoluteUrl + "/SitePages/AdmSliders.aspx";
  }

  private getQueryStringParameter(paramToRetrieve) {

    if (document.URL.indexOf("?") !== -1) {
      var params = document.URL.split("?")[1].split("&amp;");
      var strParams = "";

      for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] === paramToRetrieve)
          return singleParam[1];
      }
    }
    else return "";
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
