<%
'--------------------------------------------------------------------
' Microsoft ADO
'
' (c) 1996-1998 Microsoft Corporation.  All Rights Reserved.
'
'
'
' ADO constants include file for VBScript
'
'--------------------------------------------------------------------

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- CursorOptionEnum Values ----
Const adHoldRecords = &H00000100
Const adMovePrevious = &H00000200
Const adAddNew = &H01000400
Const adDelete = &H01000800
Const adUpdate = &H01008000
Const adBookmark = &H00002000
Const adApproxPosition = &H00004000
Const adUpdateBatch = &H00010000
Const adResync = &H00020000
Const adNotify = &H00040000
Const adFind = &H00080000
Const adSeek = &H00400000
Const adIndex = &H00800000

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- ExecuteOptionEnum Values ----
Const adRunAsync = &H00000010
Const adAsyncExecute = &H00000010
Const adAsyncFetch = &H00000020
Const adAsyncFetchNonBlocking = &H00000040
Const adExecuteNoRecords = &H00000080

'---- ConnectOptionEnum Values ----
Const adAsyncConnect = &H00000010

'---- ObjectStateEnum Values ----
Const adStateClosed = &H00000000
Const adStateOpen = &H00000001
Const adStateConnecting = &H00000002
Const adStateExecuting = &H00000004
Const adStateFetching = &H00000008

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3

'---- DataTypeEnum Values ----
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205
Const adChapter = 136
Const adFileTime = 64
Const adDBFileTime = 137
Const adPropVariant = 138
Const adVarNumeric = 139

'---- FieldAttributeEnum Values ----
Const adFldMayDefer = &H00000002
Const adFldUpdatable = &H00000004
Const adFldUnknownUpdatable = &H00000008
Const adFldFixed = &H00000010
Const adFldIsNullable = &H00000020
Const adFldMayBeNull = &H00000040
Const adFldLong = &H00000080
Const adFldRowID = &H00000100
Const adFldRowVersion = &H00000200
Const adFldCacheDeferred = &H00001000
Const adFldKeyColumn = &H00008000

'---- EditModeEnum Values ----
Const adEditNone = &H0000
Const adEditInProgress = &H0001
Const adEditAdd = &H0002
Const adEditDelete = &H0004

'---- RecordStatusEnum Values ----
Const adRecOK = &H0000000
Const adRecNew = &H0000001
Const adRecModified = &H0000002
Const adRecDeleted = &H0000004
Const adRecUnmodified = &H0000008
Const adRecInvalid = &H0000010
Const adRecMultipleChanges = &H0000040
Const adRecPendingChanges = &H0000080
Const adRecCanceled = &H0000100
Const adRecCantRelease = &H0000400
Const adRecConcurrencyViolation = &H0000800
Const adRecIntegrityViolation = &H0001000
Const adRecMaxChangesExceeded = &H0002000
Const adRecObjectOpen = &H0004000
Const adRecOutOfMemory = &H0008000
Const adRecPermissionDenied = &H0010000
Const adRecSchemaViolation = &H0020000
Const adRecDBDeleted = &H0040000

'---- GetRowsOptionEnum Values ----
Const adGetRowsRest = -1

'---- PositionEnum Values ----
Const adPosUnknown = -1
Const adPosBOF = -2
Const adPosEOF = -3

'---- enum Values ----
Const adBookmarkCurrent = 0
Const adBookmarkFirst = 1
Const adBookmarkLast = 2

'---- MarshalOptionsEnum Values ----
Const adMarshalAll = 0
Const adMarshalModifiedOnly = 1

'---- AffectEnum Values ----
Const adAffectCurrent = 1
Const adAffectGroup = 2
Const adAffectAll = 3
Const adAffectAllChapters = 4

'---- ResyncEnum Values ----
Const adResyncUnderlyingValues = 1
Const adResyncAllValues = 2

'---- CompareEnum Values ----
Const adCompareLessThan = 0
Const adCompareEqual = 1
Const adCompareGreaterThan = 2
Const adCompareNotEqual = 3
Const adCompareNotComparable = 4

'---- FilterGroupEnum Values ----
Const adFilterNone = 0
Const adFilterPendingRecords = 1
Const adFilterAffectedRecords = 2
Const adFilterFetchedRecords = 3
Const adFilterPredicate = 4
Const adFilterConflictingRecords = 5

'---- SearchDirectionEnum Values ----
Const adSearchForward = 1
Const adSearchBackward = -1

'---- PersistFormatEnum Values ----
Const adPersistADTG = 0
Const adPersistXML = 1

'---- StringFormatEnum Values ----
Const adStringXML = 0
Const adStringHTML = 1
Const adClipString = 2

'---- ConnectPromptEnum Values ----
Const adPromptAlways = 1
Const adPromptComplete = 2
Const adPromptCompleteRequired = 3
Const adPromptNever = 4

'---- ConnectModeEnum Values ----
Const adModeUnknown = 0
Const adModeRead = 1
Const adModeWrite = 2
Const adModeReadWrite = 3
Const adModeShareDenyRead = 4
Const adModeShareDenyWrite = 8
Const adModeShareExclusive = &Hc
Const adModeShareDenyNone = &H10

'---- IsolationLevelEnum Values ----
Const adXactUnspecified = &Hffffffff
Const adXactChaos = &H00000010
Const adXactReadUncommitted = &H00000100
Const adXactBrowse = &H00000100
Const adXactCursorStability = &H00001000
Const adXactReadCommitted = &H00001000
Const adXactRepeatableRead = &H00010000
Const adXactSerializable = &H00100000
Const adXactIsolated = &H00100000

'---- XactAttributeEnum Values ----
Const adXactCommitRetaining = &H00020000
Const adXactAbortRetaining = &H00040000

'---- PropertyAttributesEnum Values ----
Const adPropNotSupported = &H0000
Const adPropRequired = &H0001
Const adPropOptional = &H0002
Const adPropRead = &H0200
Const adPropWrite = &H0400

'---- ErrorValueEnum Values ----
Const adErrInvalidArgument = &Hbb9
Const adErrNoCurrentRecord = &Hbcd
Const adErrIllegalOperation = &Hc93
Const adErrInTransaction = &Hcae
Const adErrFeatureNotAvailable = &Hcb3
Const adErrItemNotFound = &Hcc1
Const adErrObjectInCollection = &Hd27
Const adErrObjectNotSet = &Hd5c
Const adErrDataConversion = &Hd5d
Const adErrObjectClosed = &He78
Const adErrObjectOpen = &He79
Const adErrProviderNotFound = &He7a
Const adErrBoundToCommand = &He7b
Const adErrInvalidParamInfo = &He7c
Const adErrInvalidConnection = &He7d
Const adErrNotReentrant = &He7e
Const adErrStillExecuting = &He7f
Const adErrOperationCancelled = &He80
Const adErrStillConnecting = &He81
Const adErrNotExecuting = &He83
Const adErrUnsafeOperation = &He84

'---- ParameterAttributesEnum Values ----
Const adParamSigned = &H0010
Const adParamNullable = &H0040
Const adParamLong = &H0080

'---- ParameterDirectionEnum Values ----
Const adParamUnknown = &H0000
Const adParamInput = &H0001
Const adParamOutput = &H0002
Const adParamInputOutput = &H0003
Const adParamReturnValue = &H0004

'---- CommandTypeEnum Values ----
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
Const adCmdFile = &H0100
Const adCmdTableDirect = &H0200

'---- EventStatusEnum Values ----
Const adStatusOK = &H0000001
Const adStatusErrorsOccurred = &H0000002
Const adStatusCantDeny = &H0000003
Const adStatusCancel = &H0000004
Const adStatusUnwantedEvent = &H0000005

'---- EventReasonEnum Values ----
Const adRsnAddNew = 1
Const adRsnDelete = 2
Const adRsnUpdate = 3
Const adRsnUndoUpdate = 4
Const adRsnUndoAddNew = 5
Const adRsnUndoDelete = 6
Const adRsnRequery = 7
Const adRsnResynch = 8
Const adRsnClose = 9
Const adRsnMove = 10
Const adRsnFirstChange = 11
Const adRsnMoveFirst = 12
Const adRsnMoveNext = 13
Const adRsnMovePrevious = 14
Const adRsnMoveLast = 15

'---- SchemaEnum Values ----
Const adSchemaProviderSpecific = -1
Const adSchemaAsserts = 0
Const adSchemaCatalogs = 1
Const adSchemaCharacterSets = 2
Const adSchemaCollations = 3
Const adSchemaColumns = 4
Const adSchemaCheckConstraints = 5
Const adSchemaConstraintColumnUsage = 6
Const adSchemaConstraintTableUsage = 7
Const adSchemaKeyColumnUsage = 8
Const adSchemaReferentialConstraints = 9
Const adSchemaTableConstraints = 10
Const adSchemaColumnsDomainUsage = 11
Const adSchemaIndexes = 12
Const adSchemaColumnPrivileges = 13
Const adSchemaTablePrivileges = 14
Const adSchemaUsagePrivileges = 15
Const adSchemaProcedures = 16
Const adSchemaSchemata = 17
Const adSchemaSQLLanguages = 18
Const adSchemaStatistics = 19
Const adSchemaTables = 20
Const adSchemaTranslations = 21
Const adSchemaProviderTypes = 22
Const adSchemaViews = 23
Const adSchemaViewColumnUsage = 24
Const adSchemaViewTableUsage = 25
Const adSchemaProcedureParameters = 26
Const adSchemaForeignKeys = 27
Const adSchemaPrimaryKeys = 28
Const adSchemaProcedureColumns = 29
Const adSchemaDBInfoKeywords = 30
Const adSchemaDBInfoLiterals = 31
Const adSchemaCubes = 32
Const adSchemaDimensions = 33
Const adSchemaHierarchies = 34
Const adSchemaLevels = 35
Const adSchemaMeasures = 36
Const adSchemaProperties = 37
Const adSchemaMembers = 38

'---- SeekEnum Values ----
Const adSeekFirstEQ = &H1
Const adSeekLastEQ = &H2
Const adSeekAfterEQ = &H4
Const adSeekAfter = &H8
Const adSeekBeforeEQ = &H10
Const adSeekBefore = &H20

'---- ADCPROP_UPDATECRITERIA_ENUM Values ----
Const adCriteriaKey = 0
Const adCriteriaAllCols = 1
Const adCriteriaUpdCols = 2
Const adCriteriaTimeStamp = 3

'---- ADCPROP_ASYNCTHREADPRIORITY_ENUM Values ----
Const adPriorityLowest = 1
Const adPriorityBelowNormal = 2
Const adPriorityNormal = 3
Const adPriorityAboveNormal = 4
Const adPriorityHighest = 5

'---- CEResyncEnum Values ----
Const adResyncNone = 0
Const adResyncAutoIncrement = 1
Const adResyncConflicts = 2
Const adResyncUpdates = 4
Const adResyncInserts = 8
Const adResyncAll = 15

'---- ADCPROP_AUTORECALC_ENUM Values ----
Const adRecalcUpFront = 0
Const adRecalcAlways = 1
%>
