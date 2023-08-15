import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPFI } from '@pnp/sp';
import styles from './GemAward.module.scss';
import { IGemAward } from './IGemAward';
import { ImageFit } from 'office-ui-fabric-react';
import { Carousel, CarouselButtonsLocation, CarouselButtonsDisplay } from '@pnp/spfx-controls-react/lib/Carousel';
import { ISiteUserInfo } from '@pnp/sp/site-users';

const GemAward = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [gemAwards, setGemAwards] = useState<IGemAward[]>([]);

	useEffect(() => {
		sp.web
			.getFolderByServerRelativePath('Gem Awards')
			.files()
			.then((gemAwardResponse) => {
				gemAwardResponse.map((fileResponse) => {
					const Email =
						fileResponse.Name.split(' ')[0].toLowerCase() +
						'.' +
						fileResponse.Name.split(' ')[1].toLowerCase() +
						'@usdtl.com';
					sp.web
						.ensureUser(Email)
						.then((response) => {
							sp.web.siteUsers
								.getByEmail(response.data.Email)()
								.then((siteUserResponse: ISiteUserInfo) => {
									sp.profiles
										.getPropertiesFor(siteUserResponse.LoginName)
										.then((profileResponse) => {
											setGemAwards((gemAwards) => [
												...gemAwards,
												{
													...fileResponse,
													Email: Email,
													imageSrc:
														'https://usdtl.sharepoint.com' + fileResponse.ServerRelativeUrl,
													title: profileResponse.DisplayName,
													showDetailsOnHover: false,
													imageFit: ImageFit.centerContain,
													PersonalUrl: profileResponse.PersonalUrl
												}
											]);
										})
										.catch((error: Error) => console.error(error.message));
								})
								.catch((error: Error) => console.error(error.message));
						})
						.catch((error: Error) => console.error(error.message));
				});
			})
			.catch((error: Error) => console.error(error.message));
	}, []);

	return (
		<div className={styles.gameAwardWp}>
			<div className={styles.heading}>
				<i className='fa fa-trophy fa-lg' aria-hidden='true' /> GEM AWARDS
			</div>
			<div className={styles.container}>
				{gemAwards.length > 0 && (
					<Carousel
						buttonsLocation={CarouselButtonsLocation.center}
						buttonsDisplay={CarouselButtonsDisplay.buttonsOnly}
						contentContainerStyles={styles.carouselContent}
						indicators={false}
						isInfinite={true}
						pauseOnHover={true}
						element={gemAwards}
						containerButtonsStyles={styles.carouselButtonsContainer}
					/>
				)}
			</div>
		</div>
	);
};

export default GemAward;
