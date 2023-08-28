//+----------------------------------------------------------------------------
//
// Datatypes
//
//-----------------------------------------------------------------------------

typedef enum TASK_DAYS {
    TASK_SUNDAY       = 0x1,
    TASK_MONDAY       = 0x2,
    TASK_TUESDAY      = 0x4,
    TASK_WEDNESDAY    = 0x8,
    TASK_THURSDAY     = 0x10,
    TASK_FRIDAY       = 0x20,
    TASK_SATURDAY     = 0x40,
} TASK_DAYS;

typedef enum TASK_WEEKS {
    TASK_FIRST_WEEK   = 1,
    TASK_SECOND_WEEK  = 2,
    TASK_THIRD_WEEK   = 3,
    TASK_FOURTH_WEEK  = 4,
    TASK_LAST_WEEK    = 5,
} TASK_WEEKS;

typedef enum TASK_MONTHS {
    TASK_JANUARY      = 0x1,
    TASK_FEBRUARY     = 0x2,
    TASK_MARCH        = 0x4,
    TASK_APRIL        = 0x8,
    TASK_MAY          = 0x10,
    TASK_JUNE         = 0x20,
    TASK_JULY         = 0x40,
    TASK_AUGUST       = 0x80,
    TASK_SEPTEMBER    = 0x100,
    TASK_OCTOBER      = 0x200,
    TASK_NOVEMBER     = 0x400,
    TASK_DECEMBER     = 0x800,
} TASK_MONTHS;

typedef enum TASK_FLAGS {
    TASK_FLAG_INTERACTIVE                  =(0x1),
    TASK_FLAG_DELETE_WHEN_DONE             =(0x2),
    TASK_FLAG_DISABLED                     =(0x4),
    TASK_FLAG_START_ONLY_IF_IDLE           =(0x10),
    TASK_FLAG_KILL_ON_IDLE_END             =(0x20),
    TASK_FLAG_DONT_START_IF_ON_BATTERIES   =(0x40),
    TASK_FLAG_KILL_IF_GOING_ON_BATTERIES   =(0x80),
    TASK_FLAG_RUN_ONLY_IF_DOCKED           =(0x100),
    TASK_FLAG_HIDDEN                       =(0x200),
    TASK_FLAG_RUN_IF_CONNECTED_TO_INTERNET =(0x400),
    TASK_FLAG_RESTART_ON_IDLE_RESUME       =(0x800),
} TASK_FLAGS;

typedef enum TASK_TRIGGER_FLAGS {
    TASK_TRIGGER_FLAG_HAS_END_DATE         =(0x1),
    TASK_TRIGGER_FLAG_KILL_AT_DURATION_END =(0x2),
    TASK_TRIGGER_FLAG_DISABLED             =(0x4),
} TASK_TRIGGER_FLAGS;

typedef enum TASK_MAXS {
    [helpstring("1440 = 60 mins/hour * 24 hrs/day since a trigger/TASK could run all day at one minute intervals.")]
    TASK_MAX_RUN_TIMES = 1440,
} TASK_MAXS;

typedef enum TASK_TRIGGER_TYPE {
    TASK_TIME_TRIGGER_ONCE            = 0,
    TASK_TIME_TRIGGER_DAILY           = 1,
    TASK_TIME_TRIGGER_WEEKLY          = 2,
    TASK_TIME_TRIGGER_MONTHLYDATE     = 3,
    TASK_TIME_TRIGGER_MONTHLYDOW      = 4,
    TASK_EVENT_TRIGGER_ON_IDLE        = 5,
    TASK_EVENT_TRIGGER_AT_SYSTEMSTART = 6,
    TASK_EVENT_TRIGGER_AT_LOGON       = 7
} TASK_TRIGGER_TYPE;

typedef struct DAILY {
    short DaysInterval;
    LONG Reserved;
} DAILY;

typedef struct WEEKLY {
    short WeeksInterval;
    short rgfDaysOfTheWeek;
    short Reserved;
} WEEKLY;

typedef struct MONTHLYDATE {
    LONG rgfDays;
    short rgfMonths;
} MONTHLYDATE;

typedef struct MONTHLYDOW {
    short wWhichWeek;
    short rgfDaysOfTheWeek;
    short rgfMonths;
} MONTHLYDOW;

typedef struct TRIGGER_TYPE_UNION {
[helpstring("It was an union but VB does not have uniouns. Use RtlMoveMemory or Lset to convert it to the needed struct. Union Members:\nDAILY Daily;\nWEEKLY Weekly;\nMONTHLYDATE MonthlyDate;\nMONTHLYDOW MonthlyDOW;")]
    short vb_compat_union[3];
} TRIGGER_TYPE_UNION;

typedef struct TASK_TRIGGER {
    [helpstring("Structure size.")]
    short cbTriggerSize;
    [helpstring("Reserved. Must be zero.")]
    short Reserved1;
    [helpstring("Trigger beginning date year.")]
    short wBeginYear;
    [helpstring("Trigger beginning date month.")]
    short wBeginMonth;
    [helpstring("Trigger beginning date day.")]
    short wBeginDay;
    [helpstring("Optional trigger ending date year.")]
    short wEndYear;
    [helpstring("Optional trigger ending date month.")]
    short wEndMonth;
    [helpstring("Optional trigger ending date day.")]
    short wEndDay;
    [helpstring("Run bracket start time hour.")]
    short wStartHour;
    [helpstring("Run bracket start time minute.")]
    short wStartMinute;
    [helpstring("Duration of run bracket.")]
    LONG MinutesDuration;
    [helpstring("Run bracket repetition interval.")]
    LONG MinutesInterval;
    [helpstring("Trigger flags.")]
    TASK_TRIGGER_FLAGS rgFlags;
    [helpstring("Trigger type.")]
    TASK_TRIGGER_TYPE TriggerType;
    [helpstring("Trigger data.")]
    TRIGGER_TYPE_UNION Type;
    [helpstring("Reserved. Must be zero.")]
    short Reserved2;
    [helpstring("Maximum number of random minutes after start time.")]
    short wRandomMinutesInterval;
} TASK_TRIGGER;

//+----------------------------------------------------------------------------
//
// Interfaces
//
//-----------------------------------------------------------------------------

//+----------------------------------------------------------------------------
//
// Interface:  ITaskTrigger
//
//-----------------------------------------------------------------------------

[
    odl,
    uuid(148BD52B-A2AB-11CE-B11F-00AA00530503),
    helpstring("Trigger object interface. A Task object may contain several of these.")
]
interface ITaskTrigger : stdole.IUnknown {

    // Methods:

    HRESULT SetTrigger(
        [in] TASK_TRIGGER *pTrigger);

    HRESULT GetTrigger(
        [in, out] TASK_TRIGGER *pTrigger);

    HRESULT GetTriggerString(
        [out, retval] LONG *ppwszTrigger);
};

//+----------------------------------------------------------------------------
//
// Interface:  IScheduledWorkItem
//
//-----------------------------------------------------------------------------
[
    odl,
    uuid(a6b952f0-a4b1-11d0-997d-00aa006887ec),
    helpstring("Abstract base class for any runnable work item that can be scheduled by the task scheduler.")
]
interface IScheduledWorkItem : stdole.IUnknown {

    // Methods concerning scheduling:

    HRESULT CreateTrigger(
        [out] short *piNewTrigger,
        [out, retval] ITaskTrigger ** ppTrigger);

    HRESULT DeleteTrigger(
        [in] short iTrigger);

    HRESULT GetTriggerCount(
        [out, retval] short *pwCount);

    HRESULT GetTrigger(
        [in] short iTrigger,
        [out, retval] ITaskTrigger ** ppTrigger);

    HRESULT GetTriggerString(
        [in] short iTrigger,
        [out, retval] LONG *ppwszTrigger);

    HRESULT GetRunTimes(
        [in] SYSTEMTIME   *pstBegin,
        [in] SYSTEMTIME   *pstEnd,
        [in, out] short   *pCount,
        [out] SYSTEMTIME  *rgstTaskTimes);

    HRESULT GetNextRunTime(
        [in, out] SYSTEMTIME *pstNextRun);

    HRESULT SetIdleWait(
        [in] short wMinutes,
        [in] short wDeadlineMinutes);

    HRESULT GetIdleWait(
        [out] short * pwMinutes,
        [out] short * pwDeadlineMinutes);

// Other methods:

    HRESULT Run(
        void);

    HRESULT Terminate(
        void);

    HRESULT EditWorkItem(
        [in] LONG hWndParent,
        [in, defaultvalue(0)] LONG dwReserved);

    HRESULT GetMostRecentRunTime(
        [out] SYSTEMTIME *pstLastRun);

    HRESULT GetStatus(
        [out, retval] LONG *phrStatus);

    HRESULT GetExitCode(
        [out, retval] LONG *pdwExitCode);

//    properties:

    HRESULT SetComment(
        [in] LPWSTR pwszComment);

    HRESULT GetComment(
        [out, retval] LONG *ppwszComment);

    HRESULT SetCreator(
        [in] LPWSTR pwszCreator);

    HRESULT GetCreator(
        [out, retval] LONG *ppwszCreator);

    HRESULT SetWorkItemData(
        [in] short cbData,
        [in] BYTE *rgbData);

    HRESULT GetWorkItemData(
        [in, out] short *pcbData,
        [in] BYTE *prgbData);

    HRESULT SetErrorRetryCount(
        [in] short wRetryCount);

    HRESULT GetErrorRetryCount(
        [out, retval] short * pwRetryCount);

    HRESULT SetErrorRetryInterval(
        [in] short wRetryInterval);

    HRESULT GetErrorRetryInterval(
        [out, retval] short * pwRetryInterval);

    HRESULT SetFlags(
        [in] LONG   dwFlags);

    HRESULT GetFlags(
        [out, retval] LONG * pdwFlags);

    HRESULT SetAccountInformation(
        [in] LPWSTR pwszAccountName,
        [in] LPWSTR pwszPassword);

    HRESULT GetAccountInformation(
        [out, retval] LONG *ppwszAccountName);
}

//+----------------------------------------------------------------------------
//
//  Interface:  ITask
//
//-----------------------------------------------------------------------------
[
    odl,
    uuid(148BD524-A2AB-11CE-B11F-00AA00530503),
    helpstring("Task object interface. The primary means of task object manipulation.")
]
interface ITask : IScheduledWorkItem {

// Properties that correspond to parameters of CreateProcess:

    HRESULT SetApplicationName(
        [in] LPWSTR  pwszApplicationName);

    HRESULT GetApplicationName(
        [out, retval] LONG *ppwszApplicationName);

    HRESULT SetParameters(
        [in] LPWSTR pwszParameters);

    HRESULT GetParameters(
        [out, retval] LONG *ppwszParameters);

    HRESULT SetWorkingDirectory(
        [in] LPWSTR pwszWorkingDirectory);

    HRESULT GetWorkingDirectory(
        [out, retval] LONG *ppwszWorkingDirectory);

    HRESULT SetPriority(
        [in] LONG   dwPriority);

    HRESULT GetPriority(
        [out, retval] LONG *pdwPriority);

// Other properties:

    HRESULT SetTaskFlags(
        [in] LONG   dwFlags);

    HRESULT GetTaskFlags(
        [out, retval] LONG *pdwFlags);

    HRESULT SetMaxRunTime(
        [in] LONG   dwMaxRunTimeMS);

    HRESULT GetMaxRunTime(
        [out, retval] LONG * pdwMaxRunTimeMS);
}

//+----------------------------------------------------------------------------
//
//  Interface:  IEnumWorkItems
//
//-----------------------------------------------------------------------------

[
    odl,
    uuid(148BD528-A2AB-11CE-B11F-00AA00530503),
    helpstring("Work item object enumerator. Enumerates the work item objects within the Tasks folder.")
]
interface IEnumWorkItems : stdole.IUnknown {

    // Methods:

    LONG Next(
        [in] LONG celt,
        [in] LONG *rgpwszNames,
        [in, out] LONG * pceltFetched);

    LONG Skip(
        [in] LONG celt);

    HRESULT Reset(void);

    HRESULT Clone(
        [out, retval] IEnumWorkItems ** ppEnumWorkItems);
}

//+----------------------------------------------------------------------------
//
//  Interface:  ISchedulingAgent
//
//-----------------------------------------------------------------------------
[
    odl,
    uuid(148BD527-A2AB-11CE-B11F-00AA00530503),
    helpstring("Scheduling agent interface. Provides location transparent manipulation of task and/or queue objects within the Tasks folder.")
]
interface ISchedulingAgent : stdole.IUnknown {

    // methods:

    HRESULT SetTargetComputer(
        [in] LPWSTR pwszComputer);

    HRESULT GetTargetComputer(
        [out, retval] LONG *ppwszComputer);

    HRESULT Enum(
        [out, retval] IEnumWorkItems ** ppEnumWorkItems);

    HRESULT Activate(
        [in] LPWSTR pwszName,
        [in] UUID * riid,
        [out, retval] ITask ** ppUnk);

    HRESULT Delete(
        [in] LPWSTR pwszName);

    HRESULT NewWorkItem(
        [in] LPWSTR pwszTaskName,
        [in] UUID * rclsid,
        [in] UUID * riid,
        [out, retval] ITask ** ppUnk);

    HRESULT AddWorkItem(
        [in] LPWSTR pwszTaskName,
        [in] IScheduledWorkItem * pWorkItem);

    HRESULT IsOfType(
        [in] LPWSTR pwszName,
        [in] UUID  * riid);
};

typedef enum TASKPAGE {
    TASKPAGE_TASK     = 0,
    TASKPAGE_SCHEDULE = 1,
    TASKPAGE_SETTINGS = 2
} TASKPAGE;

//+----------------------------------------------------------------------------
//
//  Interface:  IProvideTaskPage
//
//-----------------------------------------------------------------------------

[
    odl,
    uuid(4086658a-cbbb-11cf-b604-00c04fd8d565),
    helpstring("Task property page retrieval interface. With this interface, it is possible to retrieve one or more property pages associated with a task object. Task objects inherit this interface.")
]
interface IProvideTaskPage : stdole.IUnknown {

    // Methods:

    HRESULT GetPage(
        [in] TASKPAGE tpType,
        [in] long fPersistChanges,
        [out, retval] LONG *phPage);
}

[
    uuid(148BD520-A2AB-11CE-B11F-00AA00530503),
    helpstring("Task object")
]
coclass Task {
    [default] interface ITask;
}

[
    uuid(148BD52A-A2AB-11CE-B11F-00AA00530503),
    helpstring("Scheduling agent")
]
coclass SchedulingAgent{
    [default] interface ISchedulingAgent;
}

    typedef enum {
        SCHED_S_TASK_READY = 0x00041300,
        SCHED_S_TASK_RUNNING = 0x00041301,
        SCHED_S_TASK_DISABLED = 0x00041302,
        SCHED_S_TASK_HAS_NOT_RUN = 0x00041303,
        SCHED_S_TASK_NO_MORE_RUNS = 0x00041304,
        SCHED_S_TASK_NOT_SCHEDULED = 0x00041305,
        SCHED_S_TASK_TERMINATED = 0x00041306,
        SCHED_S_TASK_NO_VALID_TRIGGERS = 0x00041307,
        SCHED_S_EVENT_TRIGGER = 0x00041308
    } SCHED_STATUS;

    typedef enum {
        SCHED_E_TRIGGER_NOT_FOUND = 0x80041309,
        SCHED_E_TASK_NOT_READY = 0x8004130a,
        SCHED_E_TASK_NOT_RUNNING = 0x8004130b,
        SCHED_E_SERVICE_NOT_INSTALLED = 0x8004130c,
        SCHED_E_CANNOT_OPEN_TASK = 0x8004130d,
        SCHED_E_INVALID_TASK = 0x8004130e,
        SCHED_E_ACCOUNT_INFORMATION_NOT_SET = 0x8004130f,
        SCHED_E_ACCOUNT_NAME_NOT_FOUND = 0x80041310,
        SCHED_E_ACCOUNT_DBASE_CORRUPT = 0x80041311,
        SCHED_E_NO_SECURITY_SERVICES = 0x80041312,
        SCHED_E_UNKNOWN_OBJECT_VERSION = 0x80041313,
        SCHED_E_UNSUPPORTED_ACCOUNT_OPTION = 0x80041314,
        SCHED_E_SERVICE_NOT_RUNNING = 0x80041315
    } SCHED_ERRORS;