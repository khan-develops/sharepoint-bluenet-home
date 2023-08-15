import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPFI } from '@pnp/sp';
import styles from './Event.module.scss';
import * as moment from 'moment';
import { IEvent } from './IEvent';

const Event = ({ sp }: { sp: SPFI }): JSX.Element => {
	const [events, setEvents] = useState<IEvent[]>([]);

	useEffect(() => {
		sp.web.lists
			.getByTitle('Event')
			.items.select('Title', 'EventDate', 'EventDescription', 'EventLink')()
			.then((eventResponse: IEvent[]) =>
				setEvents(
					eventResponse.sort((a, b) => new Date(a.EventDate).getTime() - new Date(b.EventDate).getTime())
				)
			)
			.catch((error: Error) => console.error(error.message));
	}, []);

	return (
		<div className={styles.event}>
			<div className={styles.heading}>
				<i className='fa fa-calendar fa-lg' aria-hidden='true' /> EVENTS
			</div>
			<div className={styles.container}>
				{events.map((event, index) => (
					<div className={styles.content} key={index}>
						<div className={styles.date}>{moment(event.EventDate).format('MM/DD/YY')}</div>
						<div>
							<div className={styles.eventTitle}>{event.Title}</div>
							<div className={styles.eventDescription}>{event.EventDescription}</div>
						</div>
						<div className={styles.spacer} />
						<div className={styles.link}>
							{event.EventLink && (
								<a href={event.EventLink} target='_blank' className={styles.eventLink} rel='noreferrer'>
									<i className='fa fa-link fa-1x' aria-hidden='true' style={{ color: '#1347a4' }} />
								</a>
							)}
						</div>
					</div>
				))}
			</div>
		</div>
	);
};

export default Event;
