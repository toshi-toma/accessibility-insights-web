// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
import { has } from 'lodash';
import { autobind, IRenderFunction } from '@uifabric/utilities';
import { ActionButton } from 'office-ui-fabric-react/lib/Button';
import {
    CheckboxVisibility,
    ConstrainMode,
    DetailsList,
    IObjectWithKey,
    IDetailsRowProps,
    IColumn,
} from 'office-ui-fabric-react/lib/DetailsList';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import * as React from 'react';

import { AssessmentDefaultMessageGenerator } from '../../assessments/assessment-default-message-generator';
import {
    IGeneratedAssessmentInstance,
    IUserCapturedInstance,
    ITestStepResult,
} from '../../common/types/store-data/iassessment-result-data';
import { AssessmentInstanceTableHandler } from '../handlers/assessment-instance-table-handler';
import { ManualTestStatus } from '../../common/types/manual-test-status';
import { TestStep } from '../../assessments/types/test-step';
import { AssessmentInstanceSelectedButton } from './assessment-instance-selected-button';

export interface IAssessmentInstanceTableProps {
    instancesMap: IDictionaryStringTo<IGeneratedAssessmentInstance>;
    testStep: TestStep;
    assessmentDefaultMessageGenerator: AssessmentDefaultMessageGenerator;
}

export interface AssessmentInstanceTableItem<P = {}> extends IObjectWithKey {
    key: string;
    instance: IGeneratedAssessmentInstance<P>;
    testStepResult: ITestStepResult;
}

export interface ICapturedInstanceRowData extends IObjectWithKey {
    instance: IUserCapturedInstance;
    instanceActionButtons: JSX.Element;
}

export class AssessmentInstanceTable extends React.Component<IAssessmentInstanceTableProps> {
    private getRelevantInstanceKeys(): string[] {
        return Object.keys(this.props.instancesMap).filter(key => {
            const instance = this.props.instancesMap[key];
            return instance.testStepResults[this.props.testStep.name] != null;
        });
    }

    private getItems(): AssessmentInstanceTableItem[] {
        return this.getRelevantInstanceKeys().map(instanceKey => {
            const instance = this.props.instancesMap[instanceKey];
            return {
                key: instanceKey,
                instance: instance,
                testStepResult: instance.testStepResults[this.props.testStep.name],
            };
        });
    }

    private getColumns(): IColumn[] {
        return [
            ...this.getVisualHelperColumns(),
            ...this.getInstanceDescriptionColumns(),
            ...this.props.testStep.getInstanceStatusColumns(),
        ];
    }

    private getVisualHelperColumns(): IColumn[] {
        if (!this.props.testStep.getVisualHelperToggle) {
            return [];
        }

        return [
            {
                key: 'visualHelperColumn',
                isIconOnly: true,
                fieldName: 'this does not matter as we are using onRender function',
                minWidth: 20,
                maxWidth: 20,
                isResizable: false,
                iconName: null,
                name: null,
                ariaLabel: null,
                onColumnClick: null,
                onRender: this.onRenderVisualHelperCell,
            },
        ];
    }

    @autobind
    private onRenderVisualHelperCell(item: AssessmentInstanceTableItem): JSX.Element {
        return (
            <AssessmentInstanceSelectedButton
                test={this.props.testStep.type}
                step={this.props.testStep.name}
                selector={item.key}
                isVisualizationEnabled={item.testStepResult.isVisualizationEnabled}
                isVisible={item.testStepResult.isVisible}
                onSelected={this.actionMessageCreator.changeAssessmentVisualizationState}
            />
        );
    }

    private getInstanceDescriptionColumns(): IColumn[] {
        return this.props.testStep.columnsConfig.map(columnConfig => ({
            key: columnConfig.key,
            name: columnConfig.name,
            onRender: item => columnConfig.onRender(item.instance),
            fieldName: 'this does not matter as we are using onRender function',
            minWidth: 200,
            maxWidth: 400,
            isResizable: true,
        }));
    }

    private areAllEnabled(): boolean {
        return this.getItems().every(item => item.testStepResult.isVisualizationEnabled);
    }

    private getColumnConfigs(assessmentNavState: IAssessmentNavState, allEnabled: boolean, hasVisualHelper: boolean): IColumn[] {
        let allColumns: IColumn[] = [];
        const stepConfig = this.assessmentProvider.getStep(assessmentNavState.selectedTestType, assessmentNavState.selectedTestStep);

        if (hasVisualHelper) {
            const masterCheckbox = this.getMasterCheckboxColumn(assessmentNavState, allEnabled);
            allColumns.push(masterCheckbox);
        }

        const customColumns = this.getCustomColumns(assessmentNavState);
        allColumns = allColumns.concat(customColumns);

        const statusColumns = stepConfig.getInstanceStatusColumns();
        allColumns = allColumns.concat(statusColumns);

        return allColumns;
    }

    @autobind
    public render(): JSX.Element {
        const { testStep, instancesMap } = this.props;

        if (instancesMap == null) {
            return <Spinner className="details-view-spinner" size={SpinnerSize.large} label={'Scanning'} />;
        }

        const assessmentInstances = this.getInstanceKeys(instancesMap, assessmentNavState).map(key => {
            const instance = instancesMap[key];
            return {
                key: key,
                statusChoiceGroup: this.renderChoiceGroup(instance, key, assessmentNavState),
                visualizationButton: hasVisualHelper ? this.renderSelectedButton(instance, key, assessmentNavState) : null,
                instance: instance,
            } as IAssessmentInstanceRowData;
        });
        return assessmentInstances;

        const items: IAssessmentInstanceRowData[] = this.props.assessmentInstanceTableHandler.createAssessmentInstanceTableItems(
            instancesMap,
            testStep,
        );

        const columns: IColumn[] = this.props.assessmentInstanceTableHandler.getColumnConfigs(
            this.props.instancesMap,
            this.props.assessmentNavState,
            this.props.hasVisualHelper,
        );

        const getDefaultMessage = this.props.getDefaultMessage(this.props.assessmentDefaultMessageGenerator);
        const defaultMessageComponent = getDefaultMessage(this.props.instancesMap, this.props.assessmentNavState.selectedTestStep);

        if (defaultMessageComponent) {
            return defaultMessageComponent.message;
        }

        return (
            <div>
                {this.props.renderInstanceTableHeader(this, items)}
                <DetailsList
                    items={items}
                    columns={columns}
                    checkboxVisibility={CheckboxVisibility.hidden}
                    constrainMode={ConstrainMode.horizontalConstrained}
                    onRenderRow={this.renderRow}
                    onItemInvoked={this.onItemInvoked}
                />
            </div>
        );
    }

    @autobind
    public onItemInvoked(item: IAssessmentInstanceRowData) {
        this.updateFocusedTarget(item);
    }

    @autobind
    public renderRow(props: IDetailsRowProps, defaultRender: IRenderFunction<IDetailsRowProps>) {
        return <div onClick={() => this.updateFocusedTarget(props.item)}>{defaultRender(props)}</div>;
    }

    @autobind
    public updateFocusedTarget(item: IAssessmentInstanceRowData) {
        this.props.assessmentInstanceTableHandler.updateFocusedTarget(item.instance.target);
    }

    public renderDefaultInstanceTableHeader(items: IAssessmentInstanceRowData[]): JSX.Element {
        const disabled = !this.isAnyInstanceStatusUnknown(items, this.props.assessmentNavState.selectedTestStep);

        return (
            <ActionButton iconProps={{ iconName: 'skypeCheck' }} onClick={this.onPassUnmarkedInstances} disabled={disabled}>
                Pass unmarked instances
            </ActionButton>
        );
    }

    private isAnyInstanceStatusUnknown(items: IAssessmentInstanceRowData[], step: string): boolean {
        return items.some(
            item => has(item.instance.testStepResults, step) && item.instance.testStepResults[step].status === ManualTestStatus.UNKNOWN,
        );
    }

    @autobind
    protected onPassUnmarkedInstances(): void {
        this.props.assessmentInstanceTableHandler.passUnmarkedInstances(
            this.props.assessmentNavState.selectedTestType,
            this.props.assessmentNavState.selectedTestStep,
        );
    }
}
