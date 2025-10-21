import * as React from "react";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import {
    Persona,
    PersonaSize,
    Dropdown,
    IDropdownOption,
    PrimaryButton,
    DetailsList,
    SelectionMode,
    Stack,
    Label,
    IColumn,
    IconButton,
    NormalPeoplePicker,
    IPersonaProps,
    //   IBasePickerSuggestionsProps
} from "@fluentui/react";
import styles from "./ManageMembers.module.scss";

interface IManageMembersProps {
    context: any;
}

interface IManageMembersState {
    groups: any[];
    groupMembers: { [groupName: string]: any[] };
    selectedGroup: string;
    selectedUsersToAdd: IPersonaProps[];
    isLoading: boolean;
}

export class ManageMembers extends React.Component<IManageMembersProps, IManageMembersState> {
    private sp: SPFI;

    constructor(props: IManageMembersProps) {
        super(props);
        this.state = {
            groups: [],
            groupMembers: {},
            selectedGroup: "",
            selectedUsersToAdd: [],
            isLoading: true,
        };
        this.sp = spfi().using(SPFx(this.props.context));
    }

    async componentDidMount() {
        await this.loadGroupsAndMembers();
    }

    async loadGroupsAndMembers() {
        this.setState({ isLoading: true });
        const groups = await this.sp.web.siteGroups();
        const groupMembers: { [groupName: string]: any[] } = {};
        for (const group of groups) {
            const users = await this.sp.web.siteGroups.getById(group.Id).users();
            groupMembers[group.Title] = users;
        }
        this.setState({
            groups,
            groupMembers,
            isLoading: false,
            selectedGroup: groups.length ? groups[0].Title : "",
            selectedUsersToAdd: [],
        });
    }

    onGroupChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
        this.setState({ selectedGroup: option.key as string, selectedUsersToAdd: [] });
    };

    onRoleChange = async (user: any, newGroup: string) => {
        const { selectedGroup } = this.state;
        await this.sp.web.siteGroups.getByName(selectedGroup).users.removeById(user.Id);
        await this.sp.web.siteGroups.getByName(newGroup).users.add(user.LoginName);
        await this.loadGroupsAndMembers();
    };

    onAddUsers = async () => {
        const { selectedUsersToAdd, selectedGroup } = this.state;
        for (const user of selectedUsersToAdd) {
            await this.sp.web.siteGroups.getByName(selectedGroup).users.add(user.title as string);
        }
        this.setState({ selectedUsersToAdd: [] });
        await this.loadGroupsAndMembers();
    };

    onRemoveUser = async (user: any) => {
        const { selectedGroup } = this.state;
        await this.sp.web.siteGroups.getByName(selectedGroup).users.removeById(user.Id);
        await this.loadGroupsAndMembers();
    };

    private async onFilterChanged(filterText: string): Promise<IPersonaProps[]> {
        if (!filterText) return [];

        const users = await this.sp.web.siteUsers();

        return users.map(user => ({
            key: user.LoginName,
            text: user.Title,
            secondaryText: user.Email || user.LoginName
        })) as IPersonaProps[];
    }


    render() {
        const { groups, selectedGroup, groupMembers, selectedUsersToAdd, isLoading } = this.state;
        
        const allowedGroups = ["admin", "PM", "DM", "DH", "Finance","Business Manager","Business Unit Manager","Department Manager","Team Manager"];
        const filteredGroups = groups.filter(g => allowedGroups.includes(g.Title));
        const groupOptions = filteredGroups.map(g => ({ key: g.Title, text: g.Title }));
        const allGroupNames = filteredGroups.map(g => g.Title);

        const columns: IColumn[] = [
            {
                key: "persona",
                name: "Name",
                fieldName: "Title",
                minWidth: 150,
                onRender: user => <Persona text={user.Title} size={PersonaSize.size40} imageUrl={user.Picture}></Persona>,
            },
            {
                key: "role",
                name: "Role",
                minWidth: 120,
                onRender: user => (
                    <Dropdown
                        options={allGroupNames.map(g => ({ key: g, text: g }))}
                        selectedKey={selectedGroup}
                        onChange={(_, option) => this.onRoleChange(user, option.key as string)}
                    />
                ),
            },
            {
                key: "actions",
                name: "",
                minWidth: 50,
                onRender: (user) => (
                    <IconButton
                        iconProps={{ iconName: "Delete" }}
                        title="Remove User"
                        ariaLabel="Remove User"
                        onClick={() => this.onRemoveUser(user)}
                    />
                ),
            },
        ];

        return (
            <div className={styles.manageMembersWrapper}>
                <Label>Select Group</Label>
                <Dropdown options={groupOptions} selectedKey={selectedGroup} onChange={this.onGroupChange} style={{ width: 300 }} />

                <Stack horizontal tokens={{ childrenGap: 16 }} style={{ margin: "20px 0" }}>
                    <NormalPeoplePicker
                        onResolveSuggestions={this.onFilterChanged}
                        getTextFromItem={persona => persona.text ?? ""}
                        pickerSuggestionsProps={{
                            suggestionsHeaderText: "Suggested Users",
                            noResultsFoundText: "No users found",
                        }}
                        selectedItems={selectedUsersToAdd}
                        onChange={items => this.setState({ selectedUsersToAdd: items || [] })}
                        resolveDelay={300}
                        itemLimit={5}
                        styles={{ root: { width: 300 } }}
                    />
                    <PrimaryButton text="Add User(s)" onClick={this.onAddUsers} disabled={selectedUsersToAdd.length === 0} />
                </Stack>

                {isLoading ? (
                    <div>Loading...</div>
                ) : (
                    <DetailsList items={groupMembers[selectedGroup] || []} columns={columns} selectionMode={SelectionMode.none} />
                )}
            </div>
        );
    }
}

export default ManageMembers;
