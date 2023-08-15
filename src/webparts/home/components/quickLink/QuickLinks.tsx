import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './QuickLinks.module.scss';
import { Grid } from '@mui/material';
import { SPFI } from '@pnp/sp';
import { IQuickLinks } from './IQuickLinks';

const QuickLinks = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [quickLinks, setQuickLinks] = useState<IQuickLinks[]>([]);
    
	useEffect(() => {
		sp.web.lists
			.getByTitle('Quick Links')
			.items()
			.then((quickLinkResponse: IQuickLinks[]) => setQuickLinks(quickLinkResponse))
			.catch((error: Error) => console.error(error.message));
	}, []);

	return (
		<div className={styles.quickLinksWp}>
			<div className={styles.webpartDivHeading}>
				<i className='fa fa-link aicon' aria-hidden='true' /> QUICK LINKS
			</div>
			<Grid container spacing={1} style={{ height: '26.25em', overflowY: 'auto' }}>
				{quickLinks.map((quickLink, index) => (
					<Grid item xs={12} sm={6} md={4} lg={4} xl={4}  key={index}>
						{quickLink.LinkUrl && (
							<div className={styles.content}>
								<a href={quickLink.LinkUrl.Url} target='_blank' rel="noreferrer">
									<i className={`${quickLink.TileIcon} fa-4x`} aria-hidden='true' />
									<p className={styles.text}>{quickLink.Title}</p>
								</a>
							</div>
						)}
					</Grid>
				))}
			</Grid>
		</div>
	);
};

export default QuickLinks;
