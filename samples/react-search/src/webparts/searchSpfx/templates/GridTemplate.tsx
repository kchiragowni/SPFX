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
		let icoSrc;
		icoSrc = `${path}${icon !== null && this.unknown.indexOf(icon) === -1 ? icon : 'code'}.png`;
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

		return (
			<div className={styles.searchSpfx}>
				<div className="ms-Grid"> 
  					<div className="ms-Grid-row">

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
										<div key={index} className={(index == 0 ||index % 3 == 0) ? 'row' : '' }>
											<div className="ms-Grid-col ms-u-sm6 ms-u-md6 ms-u-lg6">
												<DocumentCard onClickHref={result.Path}>
													<DocumentCardPreview { ...this.documentPreviewProps(result.ServerRedirectedPreviewURL, this.iconUrl, result.Fileextension)} />
													<DocumentCardTitle
														title={result.Filename !== null ? result.Filename.substring(0, result.Filename.lastIndexOf('.')) : ""}
														shouldTruncate={ true }/>
													<DocumentCardActivity
														activity={this.getDateFromString(result.ModifiedOWSDATE)}
														people={
														[
															{ 
																name: this.getAuthorDisplayName(result.EditorOWSUSER), 
																profileImageSrc: result.ServerRedirectedEmbedURL,
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
																{
																	icon: 'Share',
																	onClick: (ev: any) => {
																	console.log('You clicked the pin action.');
																	ev.preventDefault();
																	ev.stopPropagation();
																	}
																}
															]
														}
														views={ 432 }
														/>
												</DocumentCard>
											</div>
										</div>);
									})
							}
						</div>
					</div>	
			</div>
		);
	}
}

const rowDiv = ({index}) => {
	
}
