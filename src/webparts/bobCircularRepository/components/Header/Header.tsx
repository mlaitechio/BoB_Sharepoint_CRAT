import * as React from 'react'
import { IHeaderProps } from './IHeaderProps'
import { IHeaderState } from './IHeaderState'
import { getResponsiveMode, Image, INavLinkGroup, INavStyles, Nav, Panel, PanelType } from '@fluentui/react';
import styles from '../BobCircularRepository.module.scss';
import { DataContext } from '../../DataContext/DataContext';
import { IBobCircularRepositoryProps } from '../IBobCircularRepositoryProps';
import { Constants } from '../../Constants/Constants';
import { IconButton } from '@fluentui/react'
import { Button, Divider, Link, Menu, MenuButton, MenuItem, MenuList, MenuPopover, MenuTrigger, Tree, TreeItem, TreeItemLayout, TreeItemPersonaLayout } from '@fluentui/react-components';
import { AddCircleRegular, AddRegular, CheckboxPersonRegular, ChevronDoubleDownRegular, ChevronUpRegular, ClockRegular, DeleteRegular, EditRegular, HomeRegular, Navigation20Regular, NavigationFilled, NavigationRegular, PhoneUpdateRegular, ShieldPersonAddRegular, TaskListLtrRegular } from '@fluentui/react-icons';

export default class Header extends React.Component<IHeaderProps, IHeaderState> {

    static contextType = DataContext;
    context!: React.ContextType<typeof DataContext>;

    private masterProps: IBobCircularRepositoryProps;

    constructor(props) {
        super(props);

        this.state = {
            openNavigationPanel: false
        }

        let providerContext = this.context;
        this.masterProps = providerContext as IBobCircularRepositoryProps;

    }

    private openUserGuide = () => {
        window.open('');
    }

    private openPublishGuide = () => {
        window.open('');
    }

    public render() {

        const { onMenuSubMenuLinkClick } = this.props
        let providerContext = this.context;
        this.masterProps = providerContext as IBobCircularRepositoryProps;
        const { userDisplayName, context, isUserChecker, isUserMaker, isUserCompliance, isUserAdmin } = this.masterProps;
        const mode = getResponsiveMode(window);
        let userProfileImg = this.generateUserPhotoLink(context.pageContext.user.email)

        let mobileMode = (mode == 0 || mode == 1);
        let tabletMode = (mode == 2);
        let desktopMode = (mode == 3 || mode == 4 || mode == 5);

        let headerClass = tabletMode ? `${styles.column7}` : `${styles.column3}  ${styles.padding}`;
        let headerImgClass = tabletMode ? `${styles.column2} ` : `${styles.column1}`;
        let headerClassTabletMode = `${styles.column6}`;
        let logoImg = context.pageContext.web.absoluteUrl + "/_api/siteiconmanager/getsitelogo";//require('../../assets/sidbilogo.png')



        // ${styles.headerBgColor}

        return (<>
            <div className={`${styles.row}  ${styles.minHeight}`}>

                {
                    tabletMode &&
                    <div className={`${styles.column1} ${styles['text-center']}`} style={{
                        display: "flex",
                        verticalAlign: "middle",
                        justifyContent: "center",
                        alignItems: "center"
                    }}>
                        <Navigation20Regular style={{ cursor: "pointer" }} onClick={() => { this.setState({ openNavigationPanel: true }) }} />
                    </div>
                }

                {
                    tabletMode && this.navigationPanel()
                }

                {(desktopMode || tabletMode) && <>

                    <div className={`${headerImgClass} ${styles.textColor} `} >

                        {/* <img alt="" src={require('../assets/TIAA.png')} /> */}
                        <Link onClick={() => { window.open(`/sites/New_intranet`, `_blank`) }}>
                            <Image src={logoImg}


                                styles={{
                                    root: { padding: tabletMode ? 0 : 5 },

                                    image: {
                                        objectFit: "contain",
                                        verticalAlign: "-webkit-baseline-middle",
                                        // minHeight: 40,
                                        height: tabletMode ? 40 : 32,
                                        width: tabletMode ? 105 : "100%"
                                        //width: responsiveMode == 5 ? 250 : `100%`
                                    }
                                }}></Image>
                        </Link>
                    </div>
                </>
                }
                {/* ${styles.textColor}  */}
                {(desktopMode || tabletMode) &&
                    <div className={`${headerClass} 
                    ${styles.headerTextAlignLeft} 
                    ${styles.minHeight} `} style={{
                            fontWeight: "var(--fontWeightSemibold)",
                            fontSize: tabletMode ? 24 : 18,
                            //borderLeft: "1px solid lightgrey",
                            //borderRight: "1px solid lightgrey"
                        }}>
                        <>
                            {`${Constants.headerText}`}
                        </>
                    </div>
                }
                {
                    (desktopMode) &&
                    <div className={`${styles.column6}`} >

                        <Button className={`${styles.formBtn}`}
                            appearance="transparent" icon={<HomeRegular />}
                            iconPosition="before"
                            onClick={() => { onMenuSubMenuLinkClick(Constants.lblHome) }}

                        >Home</Button>

                        {isUserMaker && <Menu>
                            <MenuTrigger disableButtonEnhancement>
                                <MenuButton className={`${styles.formBtn}`} appearance="transparent" icon={<TaskListLtrRegular />}>
                                    Maker
                                </MenuButton>
                            </MenuTrigger>
                            <MenuPopover>
                                <MenuList>
                                    <MenuItem className={`${styles.fontRoboto}`} icon={<AddCircleRegular />}
                                        onClick={() => { onMenuSubMenuLinkClick(Constants.lblAddCircular) }}>Add Circular</MenuItem>
                                    <MenuItem className={`${styles.fontRoboto}`}
                                        icon={<EditRegular />}
                                        onClick={() => { onMenuSubMenuLinkClick(Constants.lblEditCircular) }}>
                                        Maker Dashboard</MenuItem>
                                    {/* onClick={() => { onMenuSubMenuLinkClick(Constants.lblEditCircular) }} */}
                                    {/* <MenuItem className={`${styles.fontRoboto}`}
                                        icon={<PhoneUpdateRegular />}>
                                        Master Circular Annual Updation
                                    </MenuItem> */}
                                </MenuList>
                            </MenuPopover>
                        </Menu>}
                        {isUserCompliance && <Menu>
                            <MenuTrigger disableButtonEnhancement>
                                <MenuButton className={`${styles.formBtn}`}
                                    appearance="transparent" icon={<ShieldPersonAddRegular />}
                                    onClick={() => { }}>Compilance</MenuButton>
                            </MenuTrigger>
                            <MenuPopover>
                                <MenuList>
                                    <MenuItem className={`${styles.fontRoboto}`} icon={<ClockRegular />}
                                        onClick={() => { onMenuSubMenuLinkClick(Constants.lblPendingCompliance) }}>Pending Request</MenuItem>
                                    {/* <MenuItem className={`${styles.fontRoboto}`}
                                        icon={<PhoneUpdateRegular />}>
                                        Master Circular Annual Updation
                                    </MenuItem> */}
                                </MenuList>
                            </MenuPopover>
                        </Menu>}
                        {isUserChecker && <Menu>
                            <MenuTrigger disableButtonEnhancement>
                                <MenuButton className={`${styles.fontRoboto}`} appearance="transparent"
                                    icon={<CheckboxPersonRegular />}>Checker</MenuButton>
                            </MenuTrigger>
                            <MenuPopover>
                                <MenuList>
                                    <MenuItem
                                        onClick={() => { onMenuSubMenuLinkClick(Constants.lblPendingChecker) }}
                                        className={`${styles.fontRoboto}`}
                                        icon={<ClockRegular />}>Pending Request</MenuItem>
                                    {/* <MenuItem className={`${styles.fontRoboto}`}
                                        icon={<PhoneUpdateRegular />}>
                                        Master Circular Annual Updation
                                    </MenuItem> */}
                                </MenuList>
                            </MenuPopover>
                        </Menu>}
                    </div>
                }
                {
                    (desktopMode || tabletMode) && <>
                        <div className={`${styles.column2} ${styles.user} ${styles.textAlignEnd} `}
                            title={`${userDisplayName}`}
                            style={{}}>
                            {`${userDisplayName ?? ``}`}
                        </div>

                    </>
                }


            </div >
            <Divider appearance="subtle"></Divider>
        </>
        )
    }

    private navigationPanel = (): JSX.Element => {

        const { openNavigationPanel } = this.state;
        const { onMenuSubMenuLinkClick } = this.props;

        const { isUserChecker, isUserMaker, isUserCompliance, isUserAdmin } = this.masterProps;

        let makerLink = isUserMaker ? {
            name: 'Maker',

            //icon: `CheckList`,
            key: 'Maker',
            isExpanded: true,
            links: [
                {
                    name: 'Add Circular',
                    //url: '#',
                    onClick:() => { onMenuSubMenuLinkClick(Constants.lblAddCircular) },
                    icon: `Add`,
                    key: 'AddCircular',
                    target: '_blank',
                },
                {
                    name: 'Maker Dashboard',
                    icon: `Edit`,
                    //url: '#',
                    onClick:() => { onMenuSubMenuLinkClick(Constants.lblEditCircular) },
                    key: 'MakerDashboard',
                    target: '_blank',
                },
            ],
            target: '_blank',
        } : undefined;


        let complianceLink = isUserCompliance ? {
            name: 'Compliance',
            //url: '#',
            //icon: `Shield`,
            key: 'Compliance',
            isExpanded: true,
            links: [
                {
                    name: 'Pending Request',
                    //url: '#',
                    onClick:() => { onMenuSubMenuLinkClick(Constants.lblPendingCompliance) },
                    icon: `Clock`,
                    key: 'CmpPendingRequest',
                    target: '_blank',
                },
            ],
            target: '_blank',
        } : undefined;

        let checkerLink = isUserChecker ? {
            name: 'Checker',
            //url: '#',
            //icon: `UserFollowed`,
            key: 'Checker',
            isExpanded: true,
            links: [
                {
                    name: 'Pending Request',
                    //url: '#',
                    onClick:() => { onMenuSubMenuLinkClick(Constants.lblPendingChecker) },
                    icon: `Clock`,
                    key: 'ChkPendingRequest',
                    target: '_blank',
                },
            ],
            target: '_blank',
        } : undefined;

        const navStyles: Partial<INavStyles> = {
            root: {
                width: "auto",
                height: "auto",
                boxSizing: 'border-box',
                border: '0px solid #eee',
                overflowY: 'auto',

            },
        };

        const navLinkGroups: any[] = [
            {
                links: [
                    {
                        name: 'Home',
                        // url: `#`,
                        onClick: () => { onMenuSubMenuLinkClick(Constants.lblHome) },
                        expandAriaLabel: 'Expand Home section',
                        icon: "Home",
                        isExpanded: true,
                    },
                    makerLink ?? { key: `` },
                    complianceLink ?? { key: `` },
                    checkerLink ?? { key: `` }
                ],
            },
        ];

        let navigationJSX = <>
            < Panel
                isOpen={openNavigationPanel}
                isLightDismiss={true}
                onDismiss={() => { this.setState({ openNavigationPanel: false }) }}
                type={PanelType.smallFixedNear}

                closeButtonAriaLabel="Close"
                headerText={``}
                styles={{
                    commands: { background: "white", paddingTop: 0 },
                    headerText: {
                        fontSize: "1.3em", fontWeight: "600",
                        marginBlockStart: "0.83em", marginBlockEnd: "0.83em",
                        color: "black", fontFamily: 'Roboto'
                    },

                    main: { background: "white" },
                    content: { paddingBottom: 0, paddingLeft: 0, paddingRight: 0 },
                    navigation: {
                        borderBottom: "1px solid #ccc",
                        selectors: {
                            ".ms-Button": { color: "black", marginRight: 0 },
                            ".ms-Button:hover": { color: "black" }
                        }
                    }
                }}>
                <Nav
                    onLinkClick={() => { }}
                    selectedKey="key3"
                    ariaLabel="Nav basic example"
                    styles={navStyles}
                    groups={navLinkGroups}
                />

            </Panel>
        </>;

        return navigationJSX;
    }

    public generateUserPhotoLink(userEmail): string {
        const { context } = this.masterProps;
        const userProfilePictureLink = context.pageContext.web.absoluteUrl + "/_layouts/15/userphoto.aspx?accountname=" + encodeURIComponent(userEmail) + "&size=M";
        return userProfilePictureLink
    }
}