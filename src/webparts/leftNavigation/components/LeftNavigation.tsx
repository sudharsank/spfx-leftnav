import * as React from 'react';
import { FC, useState, useEffect } from 'react';
import styles from './LeftNavigation.module.scss';
import { SPFI } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import { INavLink, INavLinkGroup, Nav, Spinner, Stack } from 'office-ui-fabric-react';

export interface ILeftNavigationProps {
	description: string;
	isDarkTheme: boolean;
	environmentMessage: string;
	hasTeamsContext: boolean;
	userDisplayName: string;
	sp: SPFI;
}

const LeftNavigation: FC<ILeftNavigationProps> = (props) => {
	const { sp } = props;
	const [loading, setLoading] = useState<boolean>(true);
	const [menuItems, setMenuItems] = useState<INavLinkGroup[]>(undefined);
	const [selMenu, setSelMenu] = useState<string>(undefined);

	const getActiveMenuItems = async (): Promise<any[]> => {
		const filQuery = `IsActive eq 1 and Position eq 'Left'`;
		return await sp.web.lists.getByTitle('Menus').items
			.select('ID', 'Title', 'PageUrl', 'IconName', 'Sequence', 'IsParent', 'ParentMenu/Id', 'ParentMenu/Title', 'IsActive')
			.expand('ParentMenu')
			.filter(filQuery)();
	};

	const _loadLeftNavigation = async (): Promise<void> => {
		const menuItems: any[] = await getActiveMenuItems();
		if (menuItems.length > 0) {
			const navLinks: INavLinkGroup[] = [];
			const navLink: INavLink[] = [];
			if (menuItems.length > 0) {
				const fil = menuItems.filter((mi: any) => mi.IsParent);
				if (fil && fil.length > 0) {
					fil.map((item: any) => {
						const subMenus: any[] = menuItems.filter((smi: any) => !smi.IsParent && smi.ParentMenu?.Id === item.ID);
						let navsubLink: INavLink[] = [];
						if (subMenus && subMenus.length > 0) {
							subMenus.map((item: any) => {
								if (item.PageUrl?.Url.toLowerCase() === (window.location.origin + window.location.pathname).toLowerCase())
									setSelMenu(item.ID.toString());
								navsubLink.push({
									key: item.ID.toString(),
									name: item.Title,
									url: item.PageUrl?.Url,
									expandAriaLabel: item.Title,
									icon: item.IconName,
								});
							});
						}
						if (item.PageUrl?.Url.toLowerCase() === (window.location.origin + window.location.pathname).toLowerCase())
							setSelMenu(item.ID.toString());
						navLink.push({
							key: item.ID.toString(),
							name: item.Title,
							url: item.PageUrl?.Url,
							expandAriaLabel: item.Title,
							icon: item.IconName,
							links: navsubLink.length > 0 ? navsubLink : [],
							isExpanded: true
						});
						navsubLink = [];
					});
				}
				navLinks.push({ links: navLink });
				setMenuItems(navLinks);
			}
		}
		setLoading(false);
	};

	const _linkClick = (ev?: React.MouseEvent<HTMLElement, MouseEvent>, item?: INavLink): void => {
		if (item) {
			setSelMenu(item.ID.toString());
		}
	};

	useEffect(() => {
		(async () => {
			await _loadLeftNavigation();
		})();
	}, []);

	return (
		<div className={styles.leftNavigation}>
			<Stack tokens={{ childrenGap: 10 }} horizontal horizontalAlign='start'>
				<Stack.Item style={{ width: '25%', borderRight: '1px solid #CCC' }}>
					<div style={{ marginRight: '5px' }}>
						{loading ? (
							<Spinner label='Please wait...' labelPosition='top' />
						) : (
							<>
								<Nav selectedKey={selMenu} groups={menuItems} className={styles.leftNavigation} onLinkClick={_linkClick} />
							</>
						)}
					</div>
				</Stack.Item>
				<Stack.Item style={{width: '75%'}}>
					<div>
						Page Content goes here.
					</div>
				</Stack.Item>
			</Stack>
		</div>
	);
};

export default LeftNavigation;
