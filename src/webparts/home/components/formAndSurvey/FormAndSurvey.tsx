import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPFI } from '@pnp/sp';
import { IFormAndSurvey } from './IStates';
import styles from './FormAndSurvey.module.scss';
import LinkIcon from '@mui/material/Icon';
import { Grid, Chip } from '@mui/material';

const FormAndSurvey = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [formsAndSurveys, setFormsAndSurveys] = useState<IFormAndSurvey[]>([]);

	useEffect(() => {
		sp.web.lists
			.getByTitle('Forms and Surveys')
			.items.select('Title', 'DocumentLink', 'Date', 'Active')()
			.then((formAndSurveyResponse: IFormAndSurvey[]) =>
				setFormsAndSurveys(
					formAndSurveyResponse.map((formAndSurvey) => ({
						...formAndSurvey,
						Date: formAndSurvey.Date ? new Date(formAndSurvey.Date).toLocaleDateString('en-US') : null
					}))
				)
			)
			.catch((error: Error) => console.error(error.message));
	}, []);

	return (
		<div className={styles.formAndSurvey}>
			<div className={styles.heading}>
				<i className='fa fa-wpforms' aria-hidden='true' /> FORMS AND SURVEYS
			</div>
			<div className={styles.container}>
				{formsAndSurveys.map((formAndSurvey, index) => (
					<div className={styles.item} key={index}>
						{formAndSurvey.DocumentLink && formAndSurvey.Active ? (
							<Grid>
								<Chip
									className={styles.chip}
									icon={<LinkIcon />}
									component='a'
									href={formAndSurvey.DocumentLink.Url}
									clickable
									target='_blank'
									label={
										formAndSurvey.Date
											? `${formAndSurvey.Title} | Due: ${formAndSurvey.Date}`
											: formAndSurvey.Title
									}
								/>
							</Grid>
						) : (
							<Chip className={styles.chip} label={formAndSurvey.Title} />
						)}
					</div>
				))}
			</div>
		</div>
	);
};
export default FormAndSurvey;
