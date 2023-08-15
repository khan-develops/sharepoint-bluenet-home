import * as React from 'react';
import { ChangeEvent, useEffect, useState } from 'react';
import styles from './Statistics.module.scss';
import { Grid } from '@mui/material';
import { SPFI } from '@pnp/sp';
import { IStatistics } from './IStatistics';

const Statistics = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [statistics, setStatistics] = useState<IStatistics[]>([]);
	const [years, setYears] = useState<string[]>([]);
	const [selectedYear, setSelectedYear] = useState<string>('');
	useEffect(() => {
		sp.web.lists
			.getByTitle('Statistics')
			.items()
			.then((statistics: IStatistics[]) => {
                setStatistics(statistics)
                return statistics;
            })
            .then((statistics: IStatistics[]) => {
                statistics.map(statistic => {
                    const years: string[] = []
                    if(!years.find(year => year === statistic.Year)) {
                        years.push(statistic.Year)
                    }
                    setYears(years)
                    setSelectedYear(years[0])
                })
            })
			.catch((error: Error) => console.error(error.message));
	}, []);

	const _handleYearChange = (event: ChangeEvent<HTMLSelectElement>): void => {
		setSelectedYear(event.target.value);
	};
	return (
		<div className={styles.statisticsWp}>
			<div className={styles.mainHeading}>
				<i className='fa fa-bar-chart fa-lg' aria-hidden='true' /> HOW ABOUT THOSE NUMBERS
			</div>
			<div className={styles.searchDiv}>
				<select id='statistics' name='statistics' className={styles.field} onChange={_handleYearChange}>
					{years.map((year, index) => (
						<option value={year} key={index}>{year}</option>
					))}
				</select>
			</div>
			<Grid container spacing={1} className={styles.container}>
				{statistics
					.filter((statistic) => statistic.Year === selectedYear)
					.map((statistic) =>
						statistic.DocumentLink ? (
							<Grid item xs={12} sm={6} md={4} lg={4} xl={4}>
								<div className={styles.content}>
									<a href={statistic.DocumentLink.Url} target='_blank' rel='noreferrer'>
										<i className='fa fa-file-pdf-o fa-4x' aria-hidden='true' />
										<p className={styles.text}>{statistic.Title}</p>
									</a>
								</div>
							</Grid>
						) : (
							<Grid item xs={12} sm={6} md={4} lg={4} xl={4}>
								<div className={styles.contentWithoutLink}>
									<p className={styles.textWithoutLink}>{statistic.Title}</p>
								</div>
							</Grid>
						)
					)}
			</Grid>
		</div>
	);
};

export default Statistics;
