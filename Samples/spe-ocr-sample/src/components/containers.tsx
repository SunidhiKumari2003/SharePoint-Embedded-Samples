/* eslint-disable @typescript-eslint/no-unused-vars */
import React, { useEffect, useState } from 'react';
import {
    Button,
    Dialog, DialogActions, DialogContent, DialogSurface, DialogBody, DialogTitle, DialogTrigger,
    Dropdown, Option,
    Input, InputProps, InputOnChangeData,
    Label,
    Spinner,
    makeStyles, shorthands, useId
} from '@fluentui/react-components';
import type {
    OptionOnSelectData,
    SelectionEvents
} from '@fluentui/react-combobox'
import { IContainer } from "./../common/IContainer";
import SpEmbedded from '../services/spembedded';
import type { DropdownProps } from "@fluentui/react-components";
import { Files } from "./files";
import { Providers } from "@microsoft/mgt-element";
import { Container } from 'react-dom';
import { IColumn } from '../common/IColumn';
import { Stack, IStackTokens } from '@fluentui/react/lib/Stack';
import { SearchBox } from '@fluentui/react/lib/SearchBox'; 



const SpEmbeddedConst = new SpEmbedded();

const useStyles = makeStyles({
    root: {
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        ...shorthands.padding('25px'),
    },
    containerSelector: {
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        rowGap: '10px',
        ...shorthands.padding('25px'),
    },
    containerSelectorControls: {
        width: '400px',
    },
    dialogContent: {
        display: 'flex',
        flexDirection: 'column',
        rowGap: '10px',
        marginBottom: '25px'
    }
});

export const Containers = (props: any) => {
    const [containers, setContainers] = useState<IContainer[]>([]);
    const [selectedContainer, setSelectedContainer] = useState<IContainer | undefined>(undefined);
    const containerSelector = useId('containerSelector');

    const [dialogOpen, setDialogOpen] = useState(false);
    const containerName = useId('containerName');
    const [name, setName] = useState('');
    const containerDescription = useId('containerDescription');
    const [description, setDescription] = useState('');
    const [creatingContainer, setCreatingContainer] = useState(false);
    // BOOKMARK 1 - constants & hooks
    useEffect(() => {
        (async () => {
            const containers = await SpEmbeddedConst.listContainers();
            if (containers) {
                setContainers(containers);
            }
        })();
    }, []);
    const handleNameChange: InputProps["onChange"] = (event: React.ChangeEvent<HTMLInputElement>, data: InputOnChangeData): void => {
        setName(data?.value);
    };

    const handleDescriptionChange: InputProps["onChange"] = (event: React.ChangeEvent<HTMLInputElement>, data: InputOnChangeData): void => {
        setDescription(data?.value);
    };

    const onContainerCreateClick = async (event: React.MouseEvent<HTMLButtonElement>): Promise<void> => {
        setCreatingContainer(true);
        const newContainer = await SpEmbeddedConst.createContainer(name, description);

        if (newContainer) {
            setName('');
            setDescription('');
            setContainers(current => [...current, newContainer]);
            setSelectedContainer(newContainer);
            setDialogOpen(false);
        } else {
            setName('');
            setDescription('');
        }
        setCreatingContainer(false);
    }
    // BOOKMARK 2 - handlers go here
    const onContainerDropdownChange = (selectedOption: any, data: OptionOnSelectData) => {
        const selected = containers.find((container) => container.id === data.optionValue);
        setSelectedContainer(selected);
        createColumns(selected!);
    };

    const listColumns = async (container: IContainer) => {
        const graphClient = Providers.globalProvider.graph.client;
        const containerId = container.id;
        const resp = await graphClient.api(`storage/fileStorage/containers/${containerId}/columns`).version('beta').get();
        const columns = (resp.value) as IColumn[];
        console.log(columns);
    }

    const createColumns = async (container: IContainer) => {
        const graphClient = Providers.globalProvider.graph.client;
        const containerId = container.id;

        const newColumns = [{
            "name": "Merchant",
            "displayName": "Merchant",
            "description": "Name of the merchant",
            "enforceUniqueValues": false, // Must be false (true not supported with Containers)
            "hidden": false,
            "indexed": true, // Set to true to be able to search files based on this column
            "text": {// https://learn.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
                "allowMultipleLines": false,
                "appendChangesToExistingText": false,
                "linesForEditing": 0,
                "maxLength": 255
            }
        },
        {
            "name": "TransactionDate",
            "displayName": "TransactionDate",
            "description": "Date of the transaction",
            "enforceUniqueValues": false, // Must be false (true not supported with Containers)
            "hidden": false,
            "indexed": false, // Set to true to be able to search files based on this column
            "dateTime": {
                "displayAs": "friendly",
                "format": "dateOnly | dateTime"
            }
        },
        {
            "name": "Total",
            "displayName": "Total",
            "description": "Total price of the transaction",
            "enforceUniqueValues": false, // Must be false (true not supported with Containers)
            "hidden": false,
            "indexed": false, // Set to true to be able to search files based on this column
            "number": {
                "decimalPlaces": "two",
                "displayAs": "number",
                "minimum": 0
            }
        },
        {
            "name": "Currency",
            "displayName": "Currency",
            "description": "Currency of the transaction",
            "enforceUniqueValues": false, // Must be false (true not supported with Containers)
            "hidden": false,
            "indexed": true, // Set to true to be able to search files based on this column
            "text": {
                // https://learn.microsoft.com/en-us/graph/api/resources/textcolumn?view=graph-rest-1.0
                "allowMultipleLines": false,
                "appendChangesToExistingText": false,
                "linesForEditing": 0,
                "maxLength": 3
            }
        }
        ];

        newColumns.forEach(async (newColumn) => {
            //TBD: introduce wait time to avoid throttling
            try {
                const resp = await graphClient.api(`storage/fileStorage/containers/${containerId}/columns`).version('beta').post(newColumn);
                const tempColumns = (resp.value) as IColumn[];
                //TBD:error handling for auto-renaming
            } catch (error: any) {
                console.error(`Failed to create column: ${error.message}`);
            }
        });
        listColumns(container);
    }

    const getColumns = async (container: IContainer, column: IColumn) => {
        const graphClient = Providers.globalProvider.graph.client;
        const containerId = container.id;
        const columnId = column.id;
        try {
            const resp = await graphClient.api(`storage/fileStorage/containers/${containerId}/columns/${columnId}`).version('beta').get();
            console.log(resp);
        } catch (error: any) {
            console.error(`Unable to get column: ${error.message}`);
        }
    }

    const styles = useStyles();
    const stackTokens: Partial<IStackTokens> = { childrenGap: 20 };



    // BOOKMARK 3 - component rendering
    return (
        <div className={styles.root}>
            <div className={styles.containerSelector}>
                <Dropdown
                    id={containerSelector}
                    placeholder="Select a Storage Container"
                    className={styles.containerSelectorControls}
                    onOptionSelect={onContainerDropdownChange}>
                    {containers.map((option) => (
                        <Option key={option.id} value={option.id}>{option.displayName}</Option>
                    ))}
                </Dropdown>
                <Dialog open={dialogOpen} onOpenChange={(event, data) => setDialogOpen(data.open)}>

                    <DialogTrigger disableButtonEnhancement>
                        <Button className={styles.containerSelectorControls} appearance='primary'>Create a new storage Container</Button>
                    </DialogTrigger>

                    <DialogSurface>
                        <DialogBody>
                            <DialogTitle>Create a new storage Container</DialogTitle>

                            <DialogContent className={styles.dialogContent}>
                                <Label htmlFor={containerName}>Container name:</Label>
                                <Input id={containerName} className={styles.containerSelectorControls} autoFocus required
                                    value={name} onChange={handleNameChange}></Input>
                                <Label htmlFor={containerDescription}>Container description:</Label>
                                <Input id={containerDescription} className={styles.containerSelectorControls} autoFocus required
                                    value={description} onChange={handleDescriptionChange}></Input>
                                {creatingContainer &&
                                    <Spinner size='medium' label='Creating storage Container...' labelPosition='after' />
                                }
                            </DialogContent>

                            <DialogActions>
                                <DialogTrigger disableButtonEnhancement>
                                    <Button appearance="secondary" disabled={creatingContainer}>Cancel</Button>
                                </DialogTrigger>
                                <Button appearance="primary"
                                    value={name}
                                    onClick={onContainerCreateClick}
                                    disabled={creatingContainer || (name === '')}>Create storage Container</Button>
                            </DialogActions>
                        </DialogBody>
                    </DialogSurface>

                </Dialog>
            </div>
            {selectedContainer && (<Files container={selectedContainer} />)}
        </div>
    );
}

export default Containers;