import * as React from 'react';
import styles from './GraphPersona.module.scss';
import { IGraphPersonaProps } from './IGraphPersonaProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IGraphPersonaState } from './IGraphPersonaState';

import AdaptiveCard from "react-adaptivecards";
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  Persona,
  PersonaSize
} from 'office-ui-fabric-react/lib/components/Persona';

import { Link } from 'office-ui-fabric-react/lib/components/Link';

export default class GraphPersona extends React.Component<IGraphPersonaProps, IGraphPersonaState> {
  constructor(props: IGraphPersonaProps) {
    super(props);
  
    this.state = {
      name: '',
      email: '',
      phone: '',
      id: '',
      image: null
    };
  }

  public render(): React.ReactElement<IGraphPersonaProps> {
    
    return (
      <div>
        {/*<Persona primaryText={this.state.name}
                secondaryText={this.state.email}
                onRenderSecondaryText={this._renderMail}
                tertiaryText={this.state.phone}
                onRenderTertiaryText={this._renderPhone}
                optionalText={this.state.id}
                //onRenderOptionalText={this._renderID}
                imageUrl={this.state.image}
                size={PersonaSize.size100} />
    */}
        <AdaptiveCard
              payload={{
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                type: "AdaptiveCard",
                version: "1.0",
                body: [
                  {
                    type: "Container",
                    items: [
                      {
                        type: "Image",
                        horizontalAlignment: "Left",
                        spacing: "Large",
                        style: "Person",
                        url: `${
                          this.state.image
                            ? this.state.image
                            : "http://www.sharepointpals.com/image.axd?picture=/avatars/Authors/ahamed-fazil-buhari.png"
                        }`,
                        size: "Medium"
                      },
                      {
                        type: "TextBlock",
                        size: "Large",
                        horizontalAlignment: "Center",
                        weight: "Normal",
                        text: `${this.state.name}`
                      },
                      {
                        type: "TextBlock",
                        horizontalAlignment: "Center",
                        weight: "Bolder",
                        text: `${this.state.email}`,
                        wrap: true
                      }
                    ]
                  }
                ]
              }}
            />  
      </div>            
    );
  }

  private _renderMail = () => {
    if (this.state.email) {
      return <Link href={`mailto:${this.state.email}`}>{this.state.email}</Link>;
    } else {
      return <div />;
    }
  }
  
  private _renderPhone = () => {
    if (this.state.phone) {
      return <Link href={`tel:${this.state.phone}`}>{this.state.phone}</Link>;
    } else {
      return <div />;
    }
  }

  private _renderID = () => {
    if (this.state.id) {
      return <Link href={`id:${this.state.id}`}>{this.state.id}</Link>;
    } else {
      return <div />;
    }
  }  

  public componentDidMount(): void {
    this.props.graphClient
      .api('me')
      .get((error: any, user: MicrosoftGraph.User, rawResponse?: any) => {
        this.setState({
          name: user.displayName,
          email: user.mail,
          id: user.id,
          phone: user.businessPhones[0]
        });
      });
  
    this.props.graphClient
      .api('/me/photo/$value')
      .responseType('blob')
      .get((err: any, photoResponse: any, rawResponse: any) => {
        const blobUrl = window.URL.createObjectURL(photoResponse);
        this.setState({ image: blobUrl });
      });
  }    
}
