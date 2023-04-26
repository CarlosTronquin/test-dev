//loader
import { SPComponentLoader } from '@microsoft/sp-loader';

//jquery
import * as jQuery from 'jquery';

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AdmSlidersWebPart.module.scss';
import * as strings from 'AdmSlidersWebPartStrings';

//pnp
import "@pnp/sp/webs";
import "@pnp/sp/lists";

//bootstrap
SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');

export interface IAdmSlidersWebPartProps {
  description: string;
}

export default class AdmSlidersWebPart extends BaseClientSideWebPart<IAdmSlidersWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.admSliders}">
    <div class="${ styles.container}">
      <div class="${ styles.row}">
        <div class="${ styles.column}">              
          <div class="row">
            <div class="col-md-12">
              <h2>Lista sliders</h2>
              <div class="form-group">              
                <button type="button" class="btn btn-success criarSlider">Criar</button>
              </div>
            </div>
          </div>
          <div class="row">
            <div class="col-md-12">
              <div id="divSliders"></div>
            </div>
          </div>
      </div>
    </div>
  </div>`;

    var urlSite = this.context.pageContext.web.absoluteUrl;
    this.carregarSliders(urlSite);
    this.setButtonsEventHandlers();

  }

  private setButtonsEventHandlers(): void {
    const webPart: AdmSlidersWebPart = this;
    this.domElement.querySelector('button.criarSlider').addEventListener('click', () => { webPart.criarSlider(); });    
  }

  private criarSlider(): void {
    window.location.href = this.context.pageContext.web.absoluteUrl + "/SitePages/Sliders.aspx";
  }

  protected carregarSliders(urlSiteCompleta): void {

    jQuery.ajax({
      url: urlSiteCompleta + "/_api/web/lists/getbytitle('Sliders')/items?$orderby=Title asc",
      type: "GET",
      headers: {
        "Accept": "application/json;odata=verbose"
      },
      success: (data) => {
        var results = data.d.results;

        var slidersFinal = "<table class=\"table table-striped\">";
        var slidersHTML = '';
        var listItem = '';
        for (var idx = 0; idx < results.length; idx++) {
          listItem = results[idx];
          slidersHTML += "<tr><td><a target=\"_blank\" href=\" "+ listItem["Link"] + "\">" + listItem["Title"] + "</a></td>" +
            "<td><a href=\"" + urlSiteCompleta + "/SitePages/Sliders.aspx?idComunicado=" + listItem["ID"] + "\">Editar</a></td>"+
            "<td><a href=\"" + urlSiteCompleta + "/SitePages/ApagarComunicado.aspx?idComunicado=" + listItem["ID"] + "\">Excluir</a></td>";
        }

        slidersFinal += slidersHTML + "</table>";

        jQuery('#divSliders').html(slidersFinal);

      },
      error: (data) => {
        console.log("Erro ao carregar os sliders: " + data.statusText);
      }
    });
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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
