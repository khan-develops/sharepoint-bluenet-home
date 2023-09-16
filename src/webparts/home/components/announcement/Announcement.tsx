import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPFI } from '@pnp/sp';
import styles from './Announcement.module.scss';
import { Grid, Snackbar, Button, Dialog, DialogContent, DialogContentText, DialogActions, Alert } from '@mui/material';
import { ISiteUser } from '../IHome';
import { IAnnouncement } from './IAnnouncement';

const Announcement = ({ sp, currentUser }: { sp: SPFI; currentUser: ISiteUser }): JSX.Element => {
	const [announcements, setAnnouncements] = useState<IAnnouncement[]>([]);
	const [isSnackbarOpen, setIsSnackbarOpen] = useState<boolean>(false);
	const [message, setMessage] = useState<string>('Email sent successfully');
	const [severity, setSeverity] = useState<'error' | 'info' | 'success' | 'warning'>('success');
	const [isDialogOpen, setIsDialogOpen] = useState<boolean>(false);
	const [selected, setSelected] = useState<IAnnouncement[]>([]);

	const sortAnnouncement = (announcement: IAnnouncement[]): IAnnouncement[] => {
		return announcement.sort((a, b) => b.Id - a.Id);
	};

	const _getAnnouncements = (): void => {
		sp.web.lists
			.getByTitle('Announcements')
			.items()
			.then((announcementResponse) => setAnnouncements(sortAnnouncement(announcementResponse)))
			.catch((error: Error) => console.error(error.message));
	};

	useEffect(() => {
		_getAnnouncements();
	}, [selected]);

	const _handleSnackbarClose = (e?: React.SyntheticEvent, reason?: string): void => {
		if (reason === 'clickaway') {
			return;
		}
		setIsSnackbarOpen(false);
	};

	const _sendEmail = (): void => {
		sp.web.siteUserInfoList.items
			.top(5000)
			.select('EMail')()
			.then((usersResponse) => {
				usersResponse.map((user) => {
					if (user.EMail && user.EMail.split('@')[1] === 'usdtl.com') {
						const str: string = sortAnnouncement(selected)
							.map(
								(announcement) =>
									`
									<div style="padding-top:36px;padding-bottom:72px;margin-bottom:36px;border-bottom:2px solid #e1e1e1;">

										<img src=${announcement.ImageLink.DataUrl} alt="BLUENET ANNOUNCEMENT" width="auto" height="250"/>

									<div style="color:#1347a4;font-size:20px;font-weight:700;padding: 0; margin-top: 20px;">${announcement.Title}</div>
									<div style="color:#000;opacity:0.7;font-size:16px;font-weight:700;padding-bottom:12px;">Posted on ${new Date(
										announcement.Date
									).toLocaleDateString('en-US')}</div>
									<div>${announcement.Description}</div>
									</div>
									`
							)
							.join('');
						sp.utility
							.sendEmail({
								To: [user.Email],
								Subject: 'BlueNet Announcement',
								AdditionalHeaders: {
									'content-type': 'multipart/related'
								},
								Body: `
								<!DOCTYPE html>
								<html lang="en">
								<head>
								  <meta charset="utf-8">
								  <meta name="viewport" content="width=device-width,initial-scale=1">
								  <meta name="x-apple-disable-message-reformatting">
								  <title></title>
								</head>
								<body>
								  ${str}
								</body>
								</html>
								`
							})
							.then(() => {
								setMessage('success');
								setSeverity('success');
								setIsSnackbarOpen(true);
								setSelected([]);
							})
							.catch((error: Error) => {
								setMessage(error.message);
								setSeverity('error');
								setIsSnackbarOpen(true);
								setSelected([]);
							});
					}
				});
			})
			.catch((error: Error) => console.error(error.message));
	};

	const handleCheck = (event: React.ChangeEvent<HTMLInputElement>, announcement: IAnnouncement): void => {
		fetch(announcement.ImageLink.Url)
			.then((response) => response.blob())
			.then((blob) => {
				const reader = new FileReader();
				reader.readAsDataURL(blob);
				reader.onload = () => {
					announcement.ImageLink.DataUrl = reader.result;
					if (selected.length < 1) {
						setSelected((prevSelected) => [...prevSelected, announcement]);
					} else if (selected.filter((item) => item.Id === announcement.Id).length > 0) {
						const newSelected = selected.filter((item) => item.Id !== announcement.Id);
						setSelected(newSelected);
					} else {
						setSelected((prevSelected) => [...prevSelected, announcement]);
						console.log(selected);
					}
				};
			})
			.catch((error: Error) => console.error(error.message));
	};

	return (
		<div className={styles.announcementWp}>
			<div>
				<div className={styles.mainHeading}>
					<i className='fa fa-bullhorn fa-lg' aria-hidden='true' /> ANNOUNCEMENTS
					<span style={{ float: 'right' }}>
						{selected.length > 0 && (
							<Button style={{ fontSize: '0.7em' }} onClick={() => setIsDialogOpen(true)}>
								Send
							</Button>
						)}
					</span>
				</div>
			</div>
			<Snackbar open={isSnackbarOpen} autoHideDuration={6000} onClose={() => setIsSnackbarOpen(false)}>
				<Alert onClose={_handleSnackbarClose} severity={severity} style={{ fontSize: 'large' }}>
					{message}
				</Alert>
			</Snackbar>
			<Dialog open={isDialogOpen} maxWidth='md' onClose={() => setIsDialogOpen(false)}>
				<DialogContent dividers={true}>
					<DialogContentText style={{ fontSize: 'large' }}>
						Please confirm to send your email.
					</DialogContentText>
				</DialogContent>
				<DialogActions>
					<Button
						autoFocus
						style={{ fontSize: 'small' }}
						variant='outlined'
						color='primary'
						onClick={() => {
							_sendEmail();
							setIsDialogOpen(false);
						}}>
						Confirm
					</Button>
					<Button
						autoFocus
						style={{ fontSize: 'small' }}
						variant='outlined'
						color='secondary'
						onClick={() => setIsDialogOpen(false)}>
						Cancel
					</Button>
				</DialogActions>
			</Dialog>
			<div className={styles.container}>
				{sortAnnouncement(announcements).map(
					(announcement) =>
						announcement &&
						announcement.IsActive && (
							<Grid container className={styles.gridContainer} spacing={1}>
								<Grid item xs={12} sm={4} md={3} lg={3} xl={3}>
									{announcement && announcement.DocumentLink ? (
										<a
											href={announcement.DocumentLink.Url ? announcement.DocumentLink.Url : ''}
											target='_blank'
											rel='noreferrer'>
											<img
												className={styles.announcementImage}
												src={announcement.ImageLink.Url}
											/>
										</a>
									) : (
										<img className={styles.announcementImage} src={announcement.ImageLink.Url} />
									)}
								</Grid>
								<Grid item xs={12} sm={8} md={9} lg={9} xl={9}>
									<Grid container className={styles.gridContainer} spacing={1}>
										<Grid item xs={11} sm={11} md={11} lg={11} xl={11}>
											<div className={styles.announcementHeading}>{announcement.Title}</div>
											<div className={styles.announcementDate}>
												Posted on {new Date(announcement.Date).toLocaleDateString('en-US')}
											</div>
										</Grid>
										<Grid item xs={1} sm={1} md={1} lg={1} xl={1}>
											{((currentUser && currentUser.Title === 'Matt Russell') ||
												(currentUser && currentUser.Title === 'Priti Soni') ||
												(currentUser && currentUser.Title === 'Michelle Lach') ||
												(currentUser && currentUser.Title === 'Michaela Bennett') ||
												(currentUser && currentUser.Title === 'Madeline Lange') ||
												(currentUser && currentUser.Title === 'Batsaikhan Ulambayar')) && (
												<input
													style={{ float: 'right' }}
													checked={
														selected.filter((item) => item.Id === announcement.Id).length >
														0
													}
													type='checkbox'
													onChange={(e) => handleCheck(e, announcement)}
												/>
											)}
										</Grid>
										<Grid item xs={12} sm={12} md={12} lg={12} xl={12}>
											<div
												className={styles.announcementDesc}
												dangerouslySetInnerHTML={{ __html: announcement.Description }}
											/>
										</Grid>
									</Grid>
								</Grid>
							</Grid>
						)
				)}
			</div>
		</div>
	);
};

export default Announcement;
