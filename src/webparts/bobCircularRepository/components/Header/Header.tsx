import * as React from 'react'
import { IHeaderProps } from './IHeaderProps'
import { IHeaderState } from './IHeaderState'
import { getResponsiveMode, Image } from '@fluentui/react';
import styles from '../BobCircularRepository.module.scss';
import { DataContext } from '../../DataContext/DataContext';
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import { Constants } from '../../Constants/Constants';
import { IconButton } from '@fluentui/react'
import { Button, Link, Menu, MenuButton, MenuItem, MenuList, MenuPopover, MenuTrigger } from '@fluentui/react-components';
import { AddCircleRegular, AddRegular, CheckboxPersonRegular, ClockRegular, DeleteRegular, EditRegular, HomeRegular, NavigationFilled, NavigationRegular, ShieldPersonAddRegular, TaskListLtrRegular } from '@fluentui/react-icons';

export default class Header extends React.Component<IHeaderProps, IHeaderState> {

    static contextType = DataContext;
    context!: React.ContextType<typeof DataContext>;

    private masterProps: IBobCircularRepositoryProps;

    constructor(props) {
        super(props);

        this.state = {}
    }

    private openUserGuide = () => {
        window.open('');
    }

    private openPublishGuide = () => {
        window.open('');
    }

    public render() {

        const { addCircular, editCircular, deleteCircular, pendingRequest } = this.props
        let providerContext = this.context;
        this.masterProps = providerContext as IBobCircularRepositoryProps;

        const { userDisplayName, context, isUserChecker, isUserMaker, isUserCompliance, isUserAdmin } = this.masterProps;
        const mode = getResponsiveMode(window);
        let userProfileImg = this.generateUserPhotoLink(context.pageContext.user.email)

        let mobileMode = (mode == 0 || mode == 1);
        let tabletMode = (mode == 2);
        let desktopMode = (mode == 3 || mode == 4 || mode == 5);

        let headerClass = mobileMode ? `${styles.column3}` : `${styles.column3}`;
        let headerClassTabletMode = `${styles.column6}`;
        let logoImg = context.pageContext.web.absoluteUrl + "/_api/siteiconmanager/getsitelogo";//require('../../assets/sidbilogo.png')

        return (
            <div className={`${styles.row} ${styles.headerBgColor} ${styles.minHeight}`}>

                {(desktopMode) && <>

                    <div className={`${styles.column1} ${styles.textColor} `} >

                        {/* <img alt="" src={require('../assets/TIAA.png')} /> */}
                        <Image src={logoImg} styles={{
                            root: { padding: 5 },
                            image: {
                                objectFit: "contain",
                                verticalAlign: "-webkit-baseline-middle",
                                // minHeight: 40,
                                height: 50,
                                width: "90%"
                                //width: responsiveMode == 5 ? 250 : `100%`
                            }
                        }}></Image>
                    </div>
                </>
                }
                {(desktopMode) &&
                    <div className={`${headerClass} 
                    ${styles.headerTextAlignLeft} ${styles.textColor} 
                    ${styles.padding} ${styles.minHeight} `}>
                        <>
                            {`${Constants.headerText}`}
                        </>
                    </div>
                }
                {
                    (desktopMode) &&
                    <div className={`${styles.column6}`}>

                        <Button className={`${styles.formBtn}`}
                            appearance="primary" icon={<HomeRegular />}
                            iconPosition="before"
                            onClick={() => { this.props.onGoBackHome() }}

                        >Home</Button>

                        {isUserMaker && <Menu>
                            <MenuTrigger disableButtonEnhancement>
                                <MenuButton className={`${styles.formBtn}`} appearance="primary" icon={<TaskListLtrRegular />}>Circular</MenuButton>

                            </MenuTrigger>
                            <MenuPopover>
                                <MenuList>
                                    <MenuItem className={`${styles.fontRoboto}`} icon={<AddCircleRegular />}
                                        onClick={() => { this.props.addCircular() }}>Add Circular</MenuItem>
                                    <MenuItem className={`${styles.fontRoboto}`}
                                        icon={<EditRegular />}
                                        onClick={() => { this.props.editCircular() }}>
                                        Edit Circular</MenuItem>
                                    <MenuItem className={`${styles.fontRoboto}`}
                                        icon={<DeleteRegular />}
                                        onClick={() => { this.props.deleteCircular() }}>Delete Circular</MenuItem>
                                </MenuList>
                            </MenuPopover>
                        </Menu>}
                        {isUserCompliance && <Menu>
                            <MenuTrigger disableButtonEnhancement>
                                <MenuButton className={`${styles.formBtn}`}
                                    appearance="primary" icon={<ShieldPersonAddRegular /> }
                                    onClick={() => { this.props.pendingRequest() }}>Compilance</MenuButton>
                            </MenuTrigger>
                            <MenuPopover>
                                <MenuList>
                                    <MenuItem className={`${styles.fontRoboto}`} icon={<ClockRegular />}
                                        onClick={() => { this.props.pendingRequest() }}>Pending Request</MenuItem>
                                </MenuList>
                            </MenuPopover>
                        </Menu>}
                        {isUserChecker && <Menu>
                            <MenuTrigger disableButtonEnhancement>
                                <MenuButton className={`${styles.fontRoboto}`} appearance="primary" 
                                icon={<CheckboxPersonRegular />}>Checker</MenuButton>
                            </MenuTrigger>
                            <MenuPopover>
                                <MenuList>
                                    <MenuItem className={`${styles.fontRoboto}`} icon={<ClockRegular />}>Pending Request</MenuItem>
                                </MenuList>
                            </MenuPopover>
                        </Menu>}
                    </div>
                }
                {
                    (desktopMode) && <>
                        <div className={`${styles.column2} ${styles.textColor} ${styles.user} ${styles.textAlignEnd} `}
                            title={`${userDisplayName}`}
                            style={{}}>
                            {`${userDisplayName}`}
                        </div>

                    </>
                }
                {
                    tabletMode && <>

                        <div className={`${styles.column2}`}>
                            <Image src={logoImg} styles={{
                                root: { padding: 5 },
                                image: {
                                    objectFit: "contain",
                                    verticalAlign: "-webkit-baseline-middle",
                                    // minHeight: 40,
                                    height: 35,
                                    //width: responsiveMode == 5 ? 250 : `100%`
                                }
                            }}></Image>
                        </div>
                        <div className={`${headerClassTabletMode} ${styles.fontSizeTablet} ${styles.textColor} ${styles.padding} ${styles.minHeight} `}>
                            {`${Constants.headerText}`}
                        </div>
                        {/* <div className={`${styles.column1} `}>
                            <img alt="" src={`${userProfileImg}`} className={styles.userImage} />

                        </div> */}
                        <div className={`${styles.column4} ${styles.textColor} ${styles.user} ${styles.textAlignEnd} `}>
                            {`${userDisplayName}`}
                        </div>


                    </>
                }
                {(mobileMode) &&

                    <>

                        <div className={`${styles.column3}`}>
                            <Image src={logoImg} styles={{
                                root: { padding: 5 },
                                image: {
                                    objectFit: "contain",
                                    verticalAlign: "-webkit-baseline-middle",
                                    // minHeight: 40,
                                    width: "100%"
                                    //width: responsiveMode == 5 ? 250 : `100%`
                                }
                            }}></Image>
                        </div>
                        <div className={`${headerClass} ${styles.mobileFont} ${styles.textColor} ${styles.minHeight}`}>
                            {`${Constants.headerText}`}
                        </div>
                        <div className={`${styles.column3} ${styles.textColor} ${styles.verticalAlignMiddle}`}>
                            {`${userDisplayName}`}

                        </div>
                    </>
                }
            </div>
        )
    }

    public generateUserPhotoLink(userEmail): string {
        const { context } = this.masterProps;
        const userProfilePictureLink = context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?accountname=" + encodeURIComponent(userEmail) + "&size=M";
        return userProfilePictureLink
    }
}