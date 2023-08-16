import { SPComponentLoader } from '@microsoft/sp-loader';
SPComponentLoader.loadCss('https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css');
import * as React from 'react';
import styles from './Home.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Grid } from '@mui/material';
import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/sites';
import '@pnp/sp/items';
import '@pnp/sp/folders';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/files/folder';
import '@pnp/sp/lists';
import '@pnp/sp/site-users/web';
import '@pnp/sp/profiles';
import '@pnp/graph/users';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/sputilities';
import Announcement from './announcement/Announcement';
import StaffDirectory from './staffDirectory/StaffDirectory';
import QuickLinks from './quickLink/QuickLinks';
import Anniversary from './anniversary/Anniversary';
import GemAward from './gemAward/GemAward';
import { Birthday } from './birthday/Birthday';
import Statistics from './statistic/Statistics';
import Calendar from './calendar/Calendar';
import TrendingInformation from './trendingInformation/TrendingInformation';
import NewEmployees from './newEmployee/NewEmployee';
import TopBanner from './topBanner/TopBanner';
import FormAndSurvey from './formAndSurvey/FormAndSurvey';
import Event from './event/Event';
import { useEffect, useState } from 'react';
import { ISiteUser } from './IHome';

const Home = ({ context }: { context: WebPartContext }): JSX.Element => {
	const sp: SPFI = spfi().using(SPFx(context));
	const [currentUser, setCurrentUser] = useState<ISiteUser>(null);
	const [siteUsers, setSiteUsers] = useState<ISiteUser[]>([]);

	useEffect(() => {
		sp.web
			.currentUser()
			.then((currentUserResponse) => {
				sp.profiles
					.getPropertiesFor(currentUserResponse.LoginName)
					.then((userProperty) => {
						const hireDate = userProperty.UserProfileProperties.find(
							(property: { Key: string; Value: string }) => property.Key === 'SPS-HireDate'
						);
						const birthDate = userProperty.UserProfileProperties.find(
							(property: { Key: string; Value: string }) => property.Key === 'SPS-Birthday'
						);
						setCurrentUser({
							...currentUser,
							Id: currentUserResponse.Id,
							Title: userProperty.DisplayName,
							EMail: userProperty.Email,
							Name: currentUserResponse.LoginName,
							HireDate: hireDate.Value,
							BirthDate: birthDate.Value
						});
					})
					.then(() => {
						sp.web.siteUserInfoList.items
							.top(5000)
							.select('Id', 'Title', 'JobTitle', 'EMail', 'WorkPhone', 'MobilePhone', 'Office', 'Name')
							.filter('EMail ne null and FirstName ne null and LastName ne null')()
							.then(async (response: ISiteUser[]) =>
								Promise.all(
									response.map(async (user) => {
										const userProperties = await sp.profiles.getPropertiesFor(user.Name);
										if (!userProperties['odata.null']) {
											const hireDate = userProperties.UserProfileProperties.find(
												(property: { Key: string; Value: string }) =>
													property.Key === 'SPS-HireDate'
											);
											const birthDate = userProperties.UserProfileProperties.find(
												(property: { Key: string; Value: string }) =>
													property.Key === 'SPS-Birthday'
											);
											return {
												...user,
												UserUrl: userProperties.UserUrl,
												HireDate: hireDate.Value,
												BirthDate: birthDate.Value
											};
										}
									})
								)
							)
							.then((users) => {
								console.log(users);
								setSiteUsers(users.filter((user) => user));
							})
							.catch((error: Error) => console.error(error.message));
					})
					.catch((error: Error) => console.error(error.message));
			})
			.catch((error: Error) => console.error(error.message));
	}, []);

	return (
		<section className={styles.home}>
			<Grid container spacing={3}>
				<Grid item xs={12} sm={12} md={12} lg={8} xl={8}>
					<Announcement sp={sp} currentUser={currentUser} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={4}>
					<QuickLinks sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={3} xl={3}>
					<FormAndSurvey sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={6} xl={6}>
					<TopBanner sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={3} xl={3}>
					<Event sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={4}>
					<StaffDirectory sp={sp} siteUsers={siteUsers} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={4}>
					<NewEmployees sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={4}>
					<TrendingInformation sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={8}>
					<Calendar sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={4}>
					<Statistics sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={4}>
					{currentUser && (
						<Birthday
							sp={sp}
							currentUser={currentUser}
							siteUsers={siteUsers}
							setSiteUsers={setSiteUsers}
							setCurrentUser={setCurrentUser}
						/>
					)}
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={4}>
					<GemAward sp={sp} />
				</Grid>
				<Grid item xs={12} sm={12} md={12} lg={4} xl={4}>
					{currentUser && (
						<Anniversary
							sp={sp}
							currentUser={currentUser}
							siteUsers={siteUsers}
							setSiteUsers={setSiteUsers}
							setCurrentUser={setCurrentUser}
						/>
					)}
				</Grid>
			</Grid>
		</section>
	);
};

export default Home;
