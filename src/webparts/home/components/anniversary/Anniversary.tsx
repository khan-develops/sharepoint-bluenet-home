import * as React from 'react';
import styles from './Anniversary.module.scss';
import { useState, MouseEvent } from 'react';
import { SPFI } from '@pnp/sp';
import * as moment from 'moment';
import { MONTHS } from '../../common/Constants';
import { IEmailProperties } from '@pnp/sp/sputilities';
import { Button, Dialog, DialogActions, DialogContent, DialogContentText, Snackbar, Alert } from '@mui/material';
import { ISiteUser } from '../IHome';

const _getYears = (anniversaryDate: string): number => {
	if (new Date().getFullYear() - new Date(anniversaryDate).getFullYear() === 0) {
		return 1;
	}
	return new Date().getFullYear() - new Date(anniversaryDate).getFullYear() + 1;
};

const Anniversary = ({
	sp,
	currentUser,
	siteUsers,
	setSiteUsers,
	setCurrentUser
}: {
	sp: SPFI;
	currentUser: ISiteUser;
	siteUsers: ISiteUser[];
	setSiteUsers: (siteUsers: ISiteUser[]) => void;
	setCurrentUser: (currentUser: ISiteUser) => void;
}): JSX.Element => {
	const [isDialogOpen, setIsDialogOpen] = useState<boolean>(false);
	const [isSnackbarOpen, setIsSnackbarOpen] = useState<boolean>(false);
	const [message, setMessage] = useState<string>('Email sent successfully');
	const [severity, setSeverity] = useState<'error' | 'info' | 'success' | 'warning'>('success');
	const [hireDate, setHireDate] = useState<string>('');

	const updateAnniversary = (event: MouseEvent<HTMLElement>): void => {
		event.preventDefault();
		sp.profiles
			.setSingleValueProfileProperty(currentUser.Name, 'SPS-HireDate', hireDate)
			.then(() => {
				sp.web
					.currentUser()
					.then((currentUserProperty) => {
						sp.profiles
							.getPropertiesFor(currentUserProperty.LoginName)
							.then((userProperty) => {
								const hireDate = userProperty.UserProfileProperties.find(
									(property: { Key: string; Value: string }) => property.Key === 'SPS-HireDate'
								);
								setCurrentUser({
									...currentUser,
									HireDate: hireDate.Value
								});
								setSiteUsers(
									siteUsers.map((siteUser) => ({
										...siteUser,
										HireDate: siteUser.Id === currentUser.Id ? hireDate.Value : siteUser.HireDate
									}))
								);
							})
							.catch((error: Error) => console.error(error.message));
					})
					.catch((error: Error) => console.error(error.message));
			})
			.catch((error: Error) => console.error(error.message));
	};

	const sortAndFilter = (users: ISiteUser[]): ISiteUser[] => {
		return users
			.filter(
				(user) =>
					user &&
					user.HireDate &&
					user.HireDate !== '' &&
					parseInt(user.HireDate.split('/')[0]) === new Date().getMonth() + 1
			)
			.sort((userA, userB) => new Date(userA.HireDate).getTime() - new Date(userB.HireDate).getTime());
	};

	const _setEmailBody = (): string => {
		let str = ``;
		sortAndFilter(siteUsers).map(
			(user) =>
				(str =
					str +
					`  
    <tr>        
      <td style="padding-right:24px;">${user.Title}</td>
      <td style="padding-right:24px;">${moment(user.HireDate).format('MM/DD/YYYY')}</td>
      <td style="padding-right:24px;text-align:center;">${_getYears(user.HireDate)}</td>
    <tr>
    `)
		);
		return str;
	};
	const _setEmailProp = (email: string): IEmailProperties => {
		const emailProps: IEmailProperties = {
			To: [email],
			Subject: `${MONTHS[new Date().getMonth()]} ANNIVERSARIES`,
			From: 'support@usdtl.com',
			AdditionalHeaders: {
				'content-type': 'text/html'
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
      <h3>${MONTHS[new Date().getMonth()]} ANNIVERSARIES</h3>
        <table font-size:16px;text-align:left;">
          <tr style="border-bottom:2px solid #ddd;">
            <th style="padding-top:12px;padding-bottom:12px;text-align:left;">Name</th>
            <th style="padding-top:12px;padding-bottom:12px;text-align:left;">Hire Date</th>
            <th style="padding-top:12px;padding-bottom:12px;text-align:left;">Years</th>
          </tr>
          ${_setEmailBody()} 
        </table>
      </body>
      </html>
      `
		};
		return emailProps;
	};
	const _sendEmail = (): void => {
		Promise.all(
			['batsaikhan.ulambayar@usdtl.com', 'priti.soni@usdtl.com', 'matt.russell@usdtl.com'].map(
				async (email) => await sp.utility.sendEmail(_setEmailProp(email))
			)
		)
			.then(() => {
				setMessage('succes');
				setSeverity('success');
				setIsSnackbarOpen(true);
			})
			.catch((error) => {
				setMessage(error.toString());
				setSeverity('error');
				setIsSnackbarOpen(true);
			});
	};

	const _handleSnackbarClose = (e?: React.SyntheticEvent, reason?: string): void => {
		if (reason === 'clickaway') {
			return;
		}
		setIsSnackbarOpen(false);
	};

	return (
		<div className={styles.anniversaryWp}>
			<div className={styles.heading}>
				<i className='fa fa-calendar fa-lg' aria-hidden='true' /> {MONTHS[new Date().getMonth()]} ANNIVERSARIES
				{((currentUser && currentUser.EMail && currentUser.EMail === 'priti.soni@usdtl.com') ||
					(currentUser && currentUser.EMail && currentUser.EMail === 'batsaikhan.ulambayar@usdtl.com') ||
					(currentUser && currentUser.EMail && currentUser.EMail === 'matt.russell@usdtl.com')) && (
					<Button style={{ fontSize: '0.7em', float: 'right' }} onClick={() => setIsDialogOpen(true)}>
						Send {MONTHS[new Date().getMonth()]} anniversaries
					</Button>
				)}
			</div>
			<div className={styles.container}>
				{currentUser.HireDate === '' ? (
					<div className={styles.formContainer}>
						<div className={styles.anniversaryEntryRequest}>
							Oops! Looks like you have not entered your anniversary date for the Monthly Anniversary
							Celebration. Please enter your anniversary date.
						</div>
						<input
							className={styles.dateField}
							type='date'
							id='start'
							name='anniversary'
							onChange={(e) => {
								setHireDate(e.target.value);
							}}
							max={moment(new Date()).format('YYYY-MM-DD')}
						/>
						<button className={styles.submitButton} onClick={updateAnniversary} disabled={hireDate === ''}>
							Submit
						</button>
					</div>
				) : (
					<div className={styles.container}>
						{siteUsers &&
							siteUsers.length > 0 &&
							sortAndFilter(siteUsers).map((user, index) => (
								<div className={styles.content} key={index}>
									<div className={styles.day}>{user.HireDate && user.HireDate.split('/')[1]}</div>
									<div className={styles.name}>{user.Title}</div>
									<div className={styles.spacer} />
									<div className={styles.year}>
										{user.HireDate && (
											<span>
												{new Date().getFullYear() - new Date(user.HireDate).getFullYear() + 1}
											</span>
										)}
										<span> </span>
										{user.HireDate && (
											<span>
												{new Date().getFullYear() - new Date(user.HireDate).getFullYear() > 0
													? 'years'
													: 'year'}
											</span>
										)}
									</div>
								</div>
							))}
					</div>
				)}
			</div>

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

			<Snackbar open={isSnackbarOpen} autoHideDuration={6000} onClose={() => setIsSnackbarOpen(false)}>
				<Alert onClose={_handleSnackbarClose} severity={severity} style={{ fontSize: 'large' }}>
					{message}
				</Alert>
			</Snackbar>
		</div>
	);
};

export default Anniversary;
