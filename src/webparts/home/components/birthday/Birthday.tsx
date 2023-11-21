import '@pnp/sp/webs';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/lists';
import '@pnp/sp/folders';
import '@pnp/sp/profiles';
import { SPFI } from '@pnp/sp';
import * as React from 'react';
import { MouseEvent } from 'react';
import * as moment from 'moment';
import '@pnp/sp/site-users/web';
import { MONTHS } from '../../common/Constants';
import { useState } from 'react';
import styles from './IBirthday.module.scss';
import { IEmailProperties } from '@pnp/sp/sputilities';
import { ISiteUser } from '../IHome';
import {
    Alert,
    Box,
    Button,
    ButtonGroup,
    Dialog,
    DialogActions,
    DialogContent,
    DialogContentText,
    Grid,
    Paper,
    Snackbar
} from '@mui/material';

export const Birthday = ({
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
    const [isEditing, setIsEditing] = useState<boolean>(false);
    const [birthDate, setBirthDate] = useState<string>('');
    const [isDialogOpen, setIsDialogOpen] = useState<boolean>(false);
    const [isSnackbarOpen, setIsSnackbarOpen] = useState<boolean>(false);
    const [message, setMessage] = useState<string>('Email sent successfully');
    const [severity, setSeverity] = useState<'error' | 'info' | 'success' | 'warning'>('success');

    const updateBirthDate = (event: MouseEvent<HTMLElement>): void => {
        event.preventDefault();
        sp.profiles
            .setSingleValueProfileProperty(currentUser.Name, 'SPS-Birthday', birthDate)
            .then(() => {
                sp.web
                    .currentUser()
                    .then((currentUserProperty) => {
                        sp.profiles
                            .getPropertiesFor(currentUserProperty.LoginName)
                            .then((userProperty) => {
                                const birthDate = userProperty.UserProfileProperties.find(
                                    (property: { Key: string; Value: string }) =>
                                        property.Key === 'SPS-Birthday'
                                );
                                setCurrentUser({
                                    ...currentUser,
                                    BirthDate: birthDate.Value
                                });
                                setSiteUsers(
                                    siteUsers.map((siteUser) => ({
                                        ...siteUser,
                                        BirthDate:
                                            siteUser.Id === currentUser.Id
                                                ? birthDate.Value
                                                : siteUser.BirthDate
                                    }))
                                );
                                setIsEditing(false);
                            })
                            .catch((error: Error) => console.error(error.message));
                    })
                    .catch((error: Error) => console.error(error.message));
            })
            .catch((error: Error) => console.error(error.message));
    };

    const _getBirthDay = (birthDate: string): number => {
        return birthDate ? Number(birthDate.split('/')[1]) : null;
    };

    const sortAndFilter = (properties: ISiteUser[]): ISiteUser[] => {
        return properties
            .filter(
                (user) =>
                    user &&
                    user.BirthDate &&
                    user.BirthDate !== '' &&
                    parseInt(user.BirthDate.split('/')[0]) === new Date().getMonth() + 1
            )
            .sort(
                (userA, userB) =>
                    new Date(userA.BirthDate).getDate() - new Date(userB.BirthDate).getDate()
            );
    };

    const _setEmailBody = (): string => {
        let str = ``;
        sortAndFilter(siteUsers).map(
            (employee) =>
                (str =
                    str +
                    `  
    <tr>        
      <td style="padding-right:24px;">${employee.Title}</td>
      <td style="padding-right:24px;">${moment(employee.BirthDate).format('MM/DD')}</td>
    <tr>
    `)
        );
        return str;
    };
    const _setEmailProp = (email: string): IEmailProperties => {
        const emailProps: IEmailProperties = {
            To: [email],
            Subject: `${MONTHS[new Date().getMonth()]} BIRTHDAYS`,
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
      <h3>${MONTHS[new Date().getMonth()]} BIRTHDAYS</h3>
        <table font-size:16px;text-align:left;">
          <tr style="border-bottom:2px solid #ddd;">
            <th style="padding-top:12px;padding-bottom:12px;text-align:left;">Name</th>
            <th style="padding-top:12px;padding-bottom:12px;text-align:left;">Birthday</th>
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
            [
                'batsaikhan.ulambayar@usdtl.com',
                'priti.soni@usdtl.com',
                'matt.russell@usdtl.com'
            ].map(async (email) => await sp.utility.sendEmail(_setEmailProp(email)))
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
        <div className={styles.birthdayWp}>
            <div className={styles.heading}>
                <i className="fa fa-birthday-cake fa-lg" aria-hidden="true" />{' '}
                {MONTHS[new Date().getMonth()]} BIRTHDAYS
                {((currentUser &&
                    currentUser.EMail &&
                    currentUser.EMail === 'priti.soni@usdtl.com') ||
                    (currentUser &&
                        currentUser.EMail &&
                        currentUser.EMail === 'batsaikhan.ulambayar@usdtl.com') ||
                    (currentUser &&
                        currentUser.EMail &&
                        currentUser.EMail === 'matt.russell@usdtl.com')) && (
                    <Button
                        style={{ fontSize: '0.7em', float: 'right' }}
                        onClick={() => setIsDialogOpen(true)}>
                        Send {MONTHS[new Date().getMonth()]} birthdays
                    </Button>
                )}
            </div>
            <Grid container spacing={2}>
                <Grid
                    item
                    xs={12}
                    sm={12}
                    md={isEditing ? 7 : 10}
                    lg={isEditing ? 7 : 10}
                    xl={isEditing ? 7 : 10}
                    sx={{ heightMax: 430, minHeight: 300, overflowY: 'auto' }}>
                    {siteUsers &&
                        siteUsers.length > 0 &&
                        sortAndFilter(siteUsers).map((user, index) => (
                            <div className={styles.content} key={index}>
                                <div className={styles.day}>{_getBirthDay(user.BirthDate)}</div>
                                <div className={styles.name}>{user.Title}</div>
                            </div>
                        ))}
                </Grid>
                <Grid
                    display="flex"
                    justifyContent={isEditing ? 'center' : 'flex-end'}
                    alignItems="flex-start"
                    item
                    xs={12}
                    sm={12}
                    md={isEditing ? 5 : 2}
                    lg={isEditing ? 5 : 2}
                    xl={isEditing ? 5 : 2}>
                    {isEditing ? (
                        <Box
                            sx={{
                                borderLeft: '1px solid #d3d3d3',
                                padding: 2,
                                height: '100%',
                                display: 'flex',
                                flexDirection: 'column',
                                justifyContent: 'center',
                                alignItems: 'center'
                            }}>
                            <Paper
                                variant="outlined"
                                sx={{
                                    marginTop: 2,
                                    marginBottom: 2,
                                    paddingTop: 1,
                                    paddingRight: 2,
                                    paddingBottom: 1,
                                    paddingLeft: 2
                                }}>
                                <input
                                    style={{ border: 'none' }}
                                    type="date"
                                    id="start"
                                    name="birthDate"
                                    onChange={(e) => setBirthDate(e.target.value)}
                                    max={moment(new Date()).format('YYYY-MM-DD')}
                                />
                            </Paper>
                            <ButtonGroup variant="text" size="medium" sx={{ border: 'none' }}>
                                <Button onClick={updateBirthDate}>Submit</Button>
                                <Button onClick={() => setIsEditing(false)}>Cancel</Button>
                            </ButtonGroup>
                        </Box>
                    ) : (
                        <Button variant="text" onClick={() => setIsEditing(true)}>
                            Add/Edit
                        </Button>
                    )}
                </Grid>
            </Grid>

            <Dialog open={isDialogOpen} maxWidth="md" onClose={() => setIsDialogOpen(false)}>
                <DialogContent dividers={true}>
                    <DialogContentText style={{ fontSize: 'large' }}>
                        Please confirm to send your email.
                    </DialogContentText>
                </DialogContent>
                <DialogActions>
                    <Button
                        autoFocus
                        style={{ fontSize: 'small' }}
                        variant="outlined"
                        color="primary"
                        onClick={() => {
                            _sendEmail();
                            setIsDialogOpen(false);
                        }}>
                        Confirm
                    </Button>
                    <Button
                        autoFocus
                        style={{ fontSize: 'small' }}
                        variant="outlined"
                        color="secondary"
                        onClick={() => setIsDialogOpen(false)}>
                        Cancel
                    </Button>
                </DialogActions>
            </Dialog>

            <Snackbar
                open={isSnackbarOpen}
                autoHideDuration={6000}
                onClose={() => setIsSnackbarOpen(false)}>
                <Alert
                    onClose={_handleSnackbarClose}
                    severity={severity}
                    style={{ fontSize: 'large' }}>
                    {message}
                </Alert>
            </Snackbar>
        </div>
    );
};
