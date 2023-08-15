import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './TopBanner.module.scss';
import { ITopBanner } from './ITopBanner';
import { SPFI } from '@pnp/sp';
import { IFileInfo } from '@pnp/sp/files';
import { ImageFit } from '@fluentui/react/lib/Image';
import { Carousel, CarouselButtonsDisplay, CarouselButtonsLocation } from '@pnp/spfx-controls-react/lib/Carousel';

const TopBanner = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [images, setImages] = useState<ITopBanner[]>([]);

	useEffect(() => {
		sp.web
			.getFolderByServerRelativePath('Top Banner Images')
			.files()
			.then((fileResponse: IFileInfo[]) => {
				setImages(
					fileResponse.map((file) => ({
						...file,
						imageSrc: `https://usdtl.sharepoint.com/${file.ServerRelativeUrl}`,
						title: null,
						description: null,
						showDetailsOnHover: false,
						Url: `https://usdtl.sharepoint.com/${file.ServerRelativeUrl}`,
						imageFit: ImageFit.centerContain
					}))
				);
			})
			.catch((error: Error) => console.error(error.message));
	}, []);

	return (
		<div className={styles.topBanner}>
			<Carousel
				buttonsLocation={CarouselButtonsLocation.top}
				buttonsDisplay={CarouselButtonsDisplay.block}
				contentContainerStyles={styles.carouselContent}
				indicators={false}
				isInfinite={true}
				element={images}
				pauseOnHover={true}
				containerButtonsStyles={styles.carouselButtonsContainer}
			/>
		</div>
	);
};

export default TopBanner;
