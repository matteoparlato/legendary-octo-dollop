type ConflictBehaviour = "fail" | "replace" | "rename";

interface TaskInputs {
    tenantId: string;
    clientId: string;
    clientSecret: string;
    driveId: string;
    targetFolder: string;
    sourceFolder: string;
    contents: string;
    conflictBehaviour: ConflictBehaviour;
    cleanTargetFolder: boolean;
    flattenFolders: boolean;
    failOnEmptySource: boolean;
}

const TaskInputNames = {
    tenantId: "tenantId",
    clientId: 'clientId',
    clientSecret: 'clientSecret',
    driveId: 'driveId',
    targetFolder: 'targetFolder',
    sourceFolder: 'sourceFolder',
    contents: 'contents',
    conflictBehaviour: 'conflictBehaviour',
    cleanTargetFolder: 'cleanTargetFolder',
    flattenFolders: 'flattenFolders',
    failOnEmptySource: 'failOnEmptySource'
};

export {
    ConflictBehaviour,
    TaskInputs,
    TaskInputNames
};