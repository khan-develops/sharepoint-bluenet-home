import * as React from 'react';
import styles from './TrendingInformation.module.scss';
import { useEffect, useState } from 'react';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import { SPFI } from '@pnp/sp';
import { IFileInfo } from "@pnp/sp/files";

const TrendingInformation = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [files, setFiles] = useState<IFileInfo[]>([]);
	useEffect(() => {
		sp.web
			.getFolderByServerRelativePath('/Shared Documents/Misc List')
			.files()
			.then((filesResponse: IFileInfo[]) => setFiles(filesResponse))
			.catch((error: Error) => console.error(error.message));
	}, []);

	const _typeChecker = (file: IFileInfo): ApplicationType  => {
		const fileType = file.Name.split('.')[file.Name.split('.').length - 1];
		if (fileType === 'pdf') {
			return ApplicationType.PDF;
		} else if (fileType === 'docx') {
			return ApplicationType.Word;
		} else if (fileType === 'xlsx') {
			return ApplicationType.Excel;
		} else if (fileType === 'aspx') {
			return ApplicationType.ASPX;
		}
	};
	const _getName = (file: IFileInfo): string => {
		return file.Name.split('.')
			.slice(0, file.Name.split('.').length - 1)
			.join('.');
	};

	return (
		<div className={styles.trendingInformationWp}>
			<div className={styles.heading}>
				<i className='fa fa-file' aria-hidden='true' /> TRENDING INFORMATION
			</div>
			<div className={styles.container}>
				{files.map((file, index) => (
					<div className={styles.content} key={index}>
						<a className={styles.link} href={file.LinkingUri} target='_blank' rel="noreferrer">
							<FileTypeIcon
								type={IconType.image}
								application={_typeChecker(file)}
								size={ImageSize.medium}
							/>
						</a>
						<div>
							<a
								className={styles.link}
								href={
									file.LinkingUri
										? file.LinkingUri
										: `https://usdtl.sharepoint.com/${file.ServerRelativeUrl}`
								}
								target='_blank' rel="noreferrer">
								<div className={styles.title}>{_getName(file)}</div>
							</a>
							<div className={styles.date}>
								Updated on {new Date(file.TimeLastModified).toLocaleDateString('en-US')}
							</div>
						</div>
					</div>
				))}
			</div>
		</div>
	);
};

export default TrendingInformation;
