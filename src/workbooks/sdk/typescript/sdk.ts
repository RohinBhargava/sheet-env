import type {
  ForkLaunchSdkRequest,
  ForkLaunchSdkResponse,
  ForkLaunchErrorResponses,
  ForkLaunchHeaders,
} from '@forklaunch/universal-sdk';

export type CellValue = {
  address: string;
  formula?: string | null;
  number?: number | null;
  value: string
};

export type CheckResult = {
  actual: string;
  cell: string;
  expected: string;
  passed: boolean;
  sheet: string
};

export type EndpointFailure = {
  error: string
};

export type EnvState = {
  fileCount: number;
  generation: number;
  scenario: string;
  stateHash: string
};

export type FileEntry = {
  modifiedMs: number;
  path: string;
  size: number;
  version: number
};

export type FileList = {
  files: Array<FileEntry>
};

export type FileListFilesItem = FileEntry;

export type HealthResponse = {
  recordCount: number;
  routes: Array<string>;
  service: string;
  status: string
};

export type ResetInput = {
  key?: string | null;
  scenario?: string | null
};

export type SaveInput = {
  contentBase64: string;
  kind?: string | null;
  path: string
};

export type SheetCells = {
  cells: Array<CellValue>;
  path: string;
  range: string;
  sheet: string;
  sheets: Array<string>
};

export type SheetCellsCellsItem = CellValue;

export type TaskList = {
  scenario: string;
  tasks: Array<TaskSummary>
};

export type TaskListTasksItem = TaskSummary;

export type TaskSummary = {
  file: string;
  id: string;
  instructions: string;
  title: string
};

export type VerifyResult = {
  checks: Array<CheckResult>;
  passed: boolean;
  taskId: string
};

export type VerifyResultChecksItem = CheckResult;

export type VersionEntry = {
  modifiedMs: number;
  size: number;
  version: number
};

export type VersionList = {
  versions: Array<VersionEntry>
};

export type VersionListVersionsItem = VersionEntry;

export type WorkbooksApiListFilesResponseFilesItem = FileListFilesItem;

export type WorkbooksApiListTasksResponseTasksItem = TaskListTasksItem;

export type WorkbooksApiListVersionsResponseVersionsItem = VersionListVersionsItem;

export type WorkbooksApiReadCellsResponseCellsItem = SheetCellsCellsItem;

export type WorkbooksApiVerifyTaskResponseChecksItem = VerifyResultChecksItem;

export type ListFilesRequest = ForkLaunchSdkRequest<
  never,
  never,
  never,
  ForkLaunchHeaders
>;

export type ListFilesResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: FileList;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type ReadFileRequest = ForkLaunchSdkRequest<
  never,
  {
    "path": string;
    "kind"?: string | null;
    "app"?: string | null
  },
  never,
  ForkLaunchHeaders
>;

export type ReadFileResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      response: Uint8Array;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type SaveFileRequest = ForkLaunchSdkRequest<
  never,
  never,
  SaveInput,
  ForkLaunchHeaders
>;

export type SaveFileResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: FileEntry;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type DeleteFileRequest = ForkLaunchSdkRequest<
  never,
  {
    "path": string;
    "kind"?: string | null;
    "app"?: string | null
  },
  never,
  ForkLaunchHeaders
>;

export type DeleteFileResponse =
  | {
      code: 204;
      status: 204;
      contentType: '';
      response: undefined;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 204 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type ListVersionsRequest = ForkLaunchSdkRequest<
  never,
  {
    "path": string;
    "kind"?: string | null;
    "app"?: string | null
  },
  never,
  ForkLaunchHeaders
>;

export type ListVersionsResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: VersionList;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type ReadCellsRequest = ForkLaunchSdkRequest<
  never,
  {
    "path": string;
    "sheet"?: string | null;
    "range"?: string | null
  },
  never,
  ForkLaunchHeaders
>;

export type ReadCellsResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: SheetCells;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type GetEnvStateRequest = ForkLaunchSdkRequest<
  never,
  never,
  never,
  ForkLaunchHeaders
>;

export type GetEnvStateResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: EnvState;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type ResetEnvRequest = ForkLaunchSdkRequest<
  never,
  never,
  ResetInput,
  ForkLaunchHeaders
>;

export type ResetEnvResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: EnvState;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type ListTasksRequest = ForkLaunchSdkRequest<
  never,
  never,
  never,
  ForkLaunchHeaders
>;

export type ListTasksResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: TaskList;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type VerifyTaskRequest = ForkLaunchSdkRequest<
  {
    "id": string
  },
  never,
  never,
  ForkLaunchHeaders
>;

export type VerifyTaskResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: VerifyResult;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type GetHealthRequest = ForkLaunchSdkRequest<
  never,
  never,
  never,
  ForkLaunchHeaders
>;

export type GetHealthResponse =
  | {
      code: 200;
      status: 200;
      contentType: 'application/json';
      response: HealthResponse;
    }
  | ForkLaunchSdkResponse<
      Record<400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503, EndpointFailure>,
      never,
      Record<number, 'application/json'>
    >
  | ForkLaunchSdkResponse<Omit<ForkLaunchErrorResponses, 200 | 400 | 401 | 403 | 404 | 408 | 409 | 413 | 415 | 429 | 500 | 502 | 503>>;

export type WorkbooksSdkClient = {
  api: {
    listFiles: (input?: ListFilesRequest) => Promise<ListFilesResponse>;
    readFile: (input: ReadFileRequest) => Promise<ReadFileResponse>;
    saveFile: (input: SaveFileRequest) => Promise<SaveFileResponse>;
    deleteFile: (input: DeleteFileRequest) => Promise<DeleteFileResponse>;
    listVersions: (input: ListVersionsRequest) => Promise<ListVersionsResponse>;
    readCells: (input: ReadCellsRequest) => Promise<ReadCellsResponse>;
    getEnvState: (input?: GetEnvStateRequest) => Promise<GetEnvStateResponse>;
    resetEnv: (input: ResetEnvRequest) => Promise<ResetEnvResponse>;
    listTasks: (input?: ListTasksRequest) => Promise<ListTasksResponse>;
    verifyTask: (input: VerifyTaskRequest) => Promise<VerifyTaskResponse>;
  };
  health: {
    getHealth: (input?: GetHealthRequest) => Promise<GetHealthResponse>;
  };
};
