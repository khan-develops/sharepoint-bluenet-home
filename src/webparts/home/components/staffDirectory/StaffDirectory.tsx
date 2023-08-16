import * as React from 'react';
import { useState, ChangeEvent } from 'react';
import { SPFI } from '@pnp/sp';
import styles from './StaffDirectory.module.scss';
import { Grid } from '@mui/material';
import { ISiteUser } from '../IHome';

const StaffDirectory = ({ sp, siteUsers }: { sp: SPFI; siteUsers: ISiteUser[] }): JSX.Element => {
	const [searchValue, setSearchValue] = useState<string>('');

	const _handleChange = (event: ChangeEvent<HTMLInputElement>): void => {
		setSearchValue(event.target.value);
	};

	const search = (siteUsers: ISiteUser[]): ISiteUser[] => {
		const profileKeys = siteUsers[0] && Object.keys(siteUsers[0]);
		return siteUsers.filter((user) =>
			profileKeys.some(
				(profileKey: keyof ISiteUser) =>
					String(user[profileKey]).toLowerCase().indexOf(searchValue.toLowerCase()) > -1
			)
		);
	};

	return (
		<div className={styles.staffDirectory}>
			<div className={styles.mainHeading}>
				<i className='fa fa-users fa-lg' aria-hidden='true' /> STAFF DIRECTORY
			</div>
			<div className={styles.searchDiv}>
				<input className={styles.field} type='search' placeholder='Search' onChange={_handleChange} />
			</div>
			<div className={styles.innerDiv}>
				{siteUsers &&
					siteUsers.length > 0 &&
					search(siteUsers)
						.sort((a, b) => a.Id - b.Id)
						.map((profile, index) => (
							<Grid container className={styles.gridContainer} key={index}>
								<Grid item xs={12} sm={3} md={3} lg={2} xl={2} style={{ textAlign: 'center' }}>
									<a href={profile.UserUrl} target='_blank' rel='noreferrer'>
										<img
											src={`https://usdtl.sharepoint.com/_layouts/15/userphoto.aspx?size=M&username=${profile.EMail}`}
											className={styles.profileImage}
										/>
									</a>
								</Grid>
								<Grid item xs={12} sm={9} md={9} lg={10} xl={10}>
									<div>
										<div>
											<a href={profile.UserUrl} target='_blank' rel='noreferrer'>
												{profile.Title}
											</a>
										</div>
										{profile.Title && (
											<div>
												<i className='fa fa-briefcase fa-lg paddingRight' /> {profile.JobTitle}
											</div>
										)}
										<div>
											<i className='fa fa-envelope-o fa-lg paddingRight' /> {profile.EMail}
										</div>
										{profile.WorkPhone && (
											<div>
												<i className='fa fa-phone fa-lg paddingRight' /> {profile.WorkPhone}
											</div>
										)}
										{profile.MobilePhone && (
											<div>
												<i className='fa fa-mobile fa-lg paddingRight' /> {profile.MobilePhone}
											</div>
										)}
									</div>
								</Grid>
							</Grid>
						))}
			</div>
		</div>
	);
};

export default StaffDirectory;
