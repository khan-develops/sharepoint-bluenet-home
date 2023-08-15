import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPFI } from '@pnp/sp';
import { IFileInfo } from '@pnp/sp/files';
import { IElement, INewEmployee } from './INewEmployee';
import styles from './NewEmployee.module.scss';
import { ImageFit } from 'office-ui-fabric-react';
import { Carousel, CarouselButtonsLocation, CarouselButtonsDisplay } from '@pnp/spfx-controls-react/lib/Carousel';

const NewEmployees = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [elements, setElements] = useState<IElement[]>([]);

	useEffect(() => {
		sp.web
			.getFolderByServerRelativePath('New Employees Images')
			.files()
			.then((filesResponse: IFileInfo[]) => {
				const employees: INewEmployee[] = filesResponse.map((file) => ({
					Email: file.Name.split('.')[0].split(' ').join('.').toLowerCase() + '@usdtl.com',
					imageSrc: 'https://usdtl.sharepoint.com/' + file.ServerRelativeUrl,
					title: file.Name.split('.')[0],
					showDetailsOnHover: false,
					url: '',
					imageFit: ImageFit.centerContain
				}));
				return employees;
			})
			.then((filesResponse: INewEmployee[]) => {
				filesResponse.map((file) => {
					sp.web.siteUsers
						.getByEmail(file.Email)()
						.then((siteUser) => {
							sp.profiles
								.getPropertiesFor(siteUser.LoginName)
								.then((profile) =>
									setElements((elements) => [
										...elements,
										{
											title: file.title,
											imageSrc: file.imageSrc,
											showDetailsOnHover: file.showDetailsOnHover,
											imageFit: file.imageFit,
											url: profile.UserUrl,
											description: (
												<div
													style={{
														display: 'flex',
														flexDirection: 'column',
														width: '100%',
														textAlign: 'center'
													}}>
													<div>{profile.Title}</div>
													<div>
														{
															profile.UserProfileProperties.find(
																(properties: { Key: string; Value: string }) =>
																	properties.Key === 'Department'
															).Value
														}
													</div>
												</div>
											)
										}
									])
								)
								.catch((error: Error) => console.error(error.message));
						})
						.catch((error: Error) => console.error(error.message));
				});
			})
			.catch((error: Error) => console.error(error.message));
	}, []);

	return (
		<div className={styles.newEmployeeWp}>
			<div className={styles.heading}>
				<i className='fa fa-users fa-lg' aria-hidden='true' /> NEW EMPLOYEES
			</div>
			<div className={styles.container}>
				<Carousel
					buttonsLocation={CarouselButtonsLocation.top}
					buttonsDisplay={CarouselButtonsDisplay.block}
					contentContainerStyles={styles.carouselContent}
					indicators={false}
					isInfinite={true}
					element={elements}
					pauseOnHover={true}
					containerButtonsStyles={styles.carouselButtonsContainer}
				/>
			</div>
		</div>
	);
};

export default NewEmployees;
