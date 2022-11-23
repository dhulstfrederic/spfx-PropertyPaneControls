import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle

} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PropertyPaneWebPartWebPart.module.scss';
import * as strings from 'PropertyPaneWebPartWebPartStrings';

export interface IPropertyPaneWebPartWebPartProps {
  description: string;
  productname: string;
  productcost: number;
  IsCertified: boolean;
  rating: number;
  processorType: string;
  invoiceType: string;
  newProcessorType:string;
  discountCoupon: boolean;
}

export default class PropertyPaneWebPartWebPart extends BaseClientSideWebPart<IPropertyPaneWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.propertyPaneWebPart} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div> description: <strong>${escape(this.properties.description)}</strong></div>
        <div> Name: <strong>${escape(this.properties.productname)}</strong></div>
        <div> Cost: <strong>${this.properties.productcost}</strong></div>
        <div> Certified: <strong>${this.properties.IsCertified}</strong></div>
        <div> Rating: <strong>${this.properties.rating}</strong></div>
        <div> Processor: <strong>${this.properties.processorType}</strong></div>
        <div> Invoice type: <strong>${this.properties.invoiceType}</strong></div>
        <div> New Processor type: <strong>${this.properties.newProcessorType}</strong></div>
        <div> Discount coupon: <strong>${this.properties.discountCoupon}</strong></div>
      </div>
     
    </section>`;
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return new Promise<void>((resolve, reject) => {
      this.properties.productname="Mouse";
      this.properties.productcost=10;

      resolve(undefined);
    });

    return super.onInit();
  }


  protected get  disableReactivePropertyChanges(): boolean {
    // disable reactive property panel changes
   // return true;

   return false;
  }


  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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
                }),
                PropertyPaneTextField('productname', {
                  label: strings.ProductNameFieldLabel,
                  multiline: false,
                  resizable: false,
                  placeholder: "Please enter a product name", "description": "Name property field"
                }),
                PropertyPaneTextField('productcost', {
                  label: strings.ProductCostFieldLabel,
                  multiline: false,
                  resizable: false,
                  placeholder: "Please enter a product cost", "description": "Name property field"
                }),
                PropertyPaneToggle('IsCertified', {
                    key: 'isCertified',
                    label: 'Is it certified ?',
                    onText: 'certified',
                    offText: 'not certified'
                }),
                PropertyPaneSlider('rating', {
                  label:'Your rating',
                  min: 1,
                  max:10,
                  step:1,
                  showValue:true,
                  value:1
                }),
                PropertyPaneChoiceGroup('processorType', {
                  label:'Processor ',
                  options: [
                    {key:'I5', text:'Intel 5'},
                    {key:'I7', text:'Intel 7', checked: true},
                    {key:'I9', text:'Intel 9'},
                  ]
                }),                
                PropertyPaneChoiceGroup('invoiceType', {
                  label:'Invoice Type ',
                  options: [
                    {key:'MSWord', text:'Ms Word', imageSrc:'https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/word_48x1.svg',selectedImageSrc:'https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/word_48x1.svg',imageSize:{width:32,height:32}},
                    {key:'MSExcel', text:'Ms Excel', imageSrc:'https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/excel_48x1.svg',selectedImageSrc:'https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/excel_48x1.svg',imageSize:{width:32,height:32}},
                    {key:'MSPowerPoint', text:'Ms Power Point', imageSrc:'https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/powerpoint_48x1.svg',selectedImageSrc:'https://static2.sharepointonline.com/files/fabric-cdn-prod_20200430.002/assets/brand-icons/product/svg/powerpoint_48x1.svg',imageSize:{width:32,height:32}},
                  ]
                }),
                PropertyPaneDropdown('newProcessorType', {
                  label:'New Processor ',
                  options: [
                    {key:'I5', text:'Intel 5'},
                    {key:'I7', text:'Intel 7'},
                    {key:'I9', text:'Intel 9'},
                  ]
                }),                                                 
                PropertyPaneCheckbox('discountCoupon', {
                  text:'Do you have a discount coupon ',
                  checked:false,
                  disabled: false
                }),    
                PropertyPaneLink('discountCoupon', {
                  href:'http://www.amazon.fr',
                  text:'Buy from amazon',
                  target:'_BLANK',
                  popupWindowProps: {
                    height:500,
                    width:500,
                    positionWindowPosition:2,
                    title:'amazon'
                  }
                }),    

              ]
            }
          ]
        }
      ]
    };
  }
}
