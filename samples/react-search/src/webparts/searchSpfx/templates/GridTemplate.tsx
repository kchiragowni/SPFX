import * as React from 'react';
import { css } from 'office-ui-fabric-react';
import ModuleLoader from '@microsoft/sp-module-loader';

import styles from '../SearchSpfx.module.scss';
import { ISearchSpfxWebPartProps } from '../ISearchSpfxWebPartProps';

import * as moment from 'moment';

import {
  DocumentCard,
  DocumentCardActions,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps
} from 'office-ui-fabric-react/lib/DocumentCard';

import { ImageFit } from 'office-ui-fabric-react/lib/Image';

export interface IGridTemplate extends ISearchSpfxWebPartProps {
	results: any[];
}

export default class GridTemplate extends React.Component<IGridTemplate, {}> {
	private iconUrl: string = "https://spoprod-a.akamaihd.net/files/odsp-next-prod_ship-2016-08-15_20160815.002/odsp-media/images/filetypes/16/";
	private unknown: string[] = ['aspx', 'null'];

	private getAuthorDisplayName(author: string): string {
		if (author !== null) {
			const splits: string[] = author.split('|');
			return splits[1].trim();
		} else {
			return "";
		}
	}

	private getDateFromString(retrievedDate: string): string {
		if (retrievedDate !== null) {
			return moment(retrievedDate).format('DD/MM/YYYY');
		} else {
			return "";
		}
	}

	private getInitials(fullName): string {
		let initials = fullName.match(/\b\w/g) || [];
		initials = ((initials.shift() || '') + (initials.pop() || '')).toUpperCase();
		return initials;
	}
	
	private getIcoSrc(path, icon): string {
		let icoSrc = `${path}${icon !== null && this.unknown.indexOf(icon) === -1 ? icon : 'code'}.png`;
		return icoSrc;
	}

	private documentPreviewProps(PreviewImageSrc, IcoPath, IcoExtension) : IDocumentCardPreviewProps {
		 let previewProps: IDocumentCardPreviewProps = {
			previewImages: [
				{
					previewImageSrc: PreviewImageSrc,
					iconSrc: this.getIcoSrc(IcoPath, IcoExtension),
					imageFit: ImageFit.cover,
					//width: 318,
					height: 196,
					accentColor: '#ce4b1f'
				}
			],
		};

		return previewProps;
	}

	public render(): JSX.Element {
		// Load the Office UI Fabrics components css file via the module loader
    	ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.6.1/fabric.components.min.css');

		let inlineStyles = {
			docCard: {
				marginBottom: '14px'
			},
			docMeta : {
				background: '#f3f3f3'
			}
		};

		return (			
			<div className={styles.searchSpfx}>
				<div className="ms-Grid"> 

							{
								(() => {
									// Check if you need to show a title
									if (this.props.title !== "") {
										return <h1 className='ms-font-xxl'>{this.props.title}</h1>;
									}
								})()
							}
							
							{
								this.props.results.map((result, index) => {
									return (									
										<div key={index}>
											<div className="ms-Grid-col ms-u-sm6 ms-u-md4 ms-u-lg4" style={inlineStyles.docCard}>
												<DocumentCard>
													<DocumentCardPreview { ...this.documentPreviewProps(result.ServerRedirectedPreviewURL, this.iconUrl, result.Fileextension)} />
													<div style={inlineStyles.docMeta}>
													<DocumentCardTitle
														title={result.Title !== null ? result.Title : ""}
														shouldTruncate={ true }/>
													<DocumentCardActivity
														activity={this.getDateFromString(result.ModifiedOWSDATE)}
														people={
														[
															{ 
																name: this.getAuthorDisplayName(result.EditorOWSUSER), 
																profileImageSrc: result.ows_PictureURL,
																initials: this.getInitials(this.getAuthorDisplayName(result.EditorOWSUSER))
															}
														]
														}
													/>
													<DocumentCardActions
														actions={
															[
																{
																	icon: 'Add',
																	onClick: (ev: any) => {
																	console.log('You clicked the share action.');
																	ev.preventDefault();
																	ev.stopPropagation();
																	}
																},
															/*	{
																	icon: 'Share',
																	onClick: (ev: any) => {
																	console.log('You clicked the pin action.');
																	ev.preventDefault();
																	ev.stopPropagation();
																	}
																},*/
																{
																	icon: 'Edit',
																	onClick: (ev: any) => {
																		location.href = (result.ServerRedirectedURL) ? result.ServerRedirectedURL : result.ServerRedirectedEmbedURL;
																	}
																}
															]
														}
														views={ result.ViewsRecent }
														/>
													</div>
												</DocumentCard>
											</div>											
										</div>);
									
									})
							}
						
					</div>	
			</div>
		);
	}
}
