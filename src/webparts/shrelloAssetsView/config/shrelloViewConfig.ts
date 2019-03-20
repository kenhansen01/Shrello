import { IconNames } from "office-ui-fabric-react/lib/Icons";
import { ColorClassNames, IDatePickerStrings } from "office-ui-fabric-react";

export interface IShrelloLabel {
  name: string;
  iconName: IconNames;
  iconColor: string;
  text?: string;
}

const ListTitles: {
  SupportDepartments: string;
  SupportTopics: string;
  ShrelloItems: string;
} = {
  SupportDepartments: "Support Departments",
  SupportTopics: "Support Topics",
  ShrelloItems: "TASC Tickets"
};

const Priorities: IShrelloLabel[] = [
  {
    name: "1 - Top",
    iconName: IconNames.FlameSolid,
    iconColor: ColorClassNames.red
  },
  {
    name: "2 - High",
    iconName: IconNames.CircleFill,
    iconColor: ColorClassNames.orange
  },
  {
    name: "3 - Normal",
    iconName: IconNames.CircleHalfFull,
    iconColor: ColorClassNames.green
  },
  {
    name: "4 - Low",
    iconName: IconNames.CircleRing,
    iconColor: ColorClassNames.blue
  }
];

const Categories: IShrelloLabel[] = [
  {
    name: "Unknown",
    iconName: IconNames.Unknown,
    iconColor: ColorClassNames.black
  },
  {
    name: "Access Request",
    iconName: IconNames.Permissions,
    iconColor: ColorClassNames.orange
  },
  {
    name: "Bug",
    iconName: IconNames.BugSolid,
    iconColor: ColorClassNames.red
  },
  {
    name: "Enhancement",
    iconName: IconNames.Lightbulb,
    iconColor: ColorClassNames.yellow
  },
  {
    name: "Suggestion",
    iconName: IconNames.Comment,
    iconColor: ColorClassNames.tealDark
  },
  {
    name: "Data Issue",
    iconName: IconNames.Database,
    iconColor: ColorClassNames.blue
  },
  {
    name: "Service Request",
    iconName: IconNames.IssueSolid,
    iconColor: ColorClassNames.magenta
  },
  {
    name: "Other Request",
    iconName: IconNames.TaskLogo,
    iconColor: ColorClassNames.green
  },
];

const Statuses: IShrelloLabel[] = [
  {
    name: "Initiated",
    iconName: IconNames.CircleAddition,
    iconColor: ColorClassNames.greenDark
  },
  {
    name: "Assigned",
    iconName: IconNames.Backlog,
    iconColor: ColorClassNames.green,
  },
  {
    name: "Planned",
    iconName: IconNames.Sprint,
    iconColor: ColorClassNames.blueMid
  },
  {
    name: "In Progress",
    iconName: IconNames.WorkItem,
    iconColor: ColorClassNames.orange
  },
  {
    name: "Ready for Review",
    iconName: IconNames.ReviewSolid,
    iconColor: ColorClassNames.tealDark
  },
  {
    name: "Ready for Release",
    iconName: IconNames.Clock,
    iconColor: ColorClassNames.magentaDark
  },
  {
    name: "On Hold",
    iconName: IconNames.CirclePause,
    iconColor: ColorClassNames.red
  },
  {
    name: "Completed",
    iconName: IconNames.Completed,
    iconColor: ColorClassNames.orangeLight
  },
  {
    name: "Canceled",
    iconName: IconNames.Cancel,
    iconColor: ColorClassNames.black
  }
];

const DayPickerStrings: IDatePickerStrings = {
  months: [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
  ],

  shortMonths: [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec"
  ],

  days: [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday"
  ],

  shortDays: [
    "S",
    "M",
    "T",
    "W",
    "T",
    "F",
    "S"
  ],

  goToToday: "Go to today",
  prevMonthAriaLabel: "Go to previous month",
  nextMonthAriaLabel: "Go to next month",
  prevYearAriaLabel: "Go to previous year",
  nextYearAriaLabel: "Go to next year",

  isRequiredErrorMessage: "Start date is required.",

  invalidInputErrorMessage: "Invalid date format."
};

export {
  ListTitles,
  Priorities,
  DayPickerStrings,
  Categories,
  Statuses
};