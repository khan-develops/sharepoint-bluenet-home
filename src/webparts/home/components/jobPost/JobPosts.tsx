import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './JobPosts.module.scss';
import { IFile } from './IJobPosts';
import { SPFI } from '@pnp/sp';
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';

const JobPosts = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [files, setFiles] = useState<IFile[]>([]);

	useEffect(() => {
		sp.web
			.getFolderByServerRelativePath('/Shared Documents/Job Posts')
			.files()
			.then((filesResponse: IFile[]) => setFiles(filesResponse))
			.catch((error: Error) => console.error(error.message));
	}, []);

	const _getName = (file: IFile): string => {
		return file.Name.split('.')
			.slice(0, file.Name.split('.').length - 1)
			.join('.');
	};

	const _typeChecker = (file: IFile): ApplicationType => {
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
	return (
		<div className={styles.jobPostsWp}>
			<div className={styles.heading}>
				<i className='fa fa-hacker-news fa-lg' aria-hidden='true' /> JOB POSTINGS
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

export default JobPosts;
