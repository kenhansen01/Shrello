import { IWebPartContext } from "@microsoft/sp-webpart-base";
import {
  IShrelloItem,
  ItemCreationCallback,
  ISupportDepartmentItem,
  ISupportTopicItem,
  ItemOperationCallback
} from "../../models";


interface IShrelloFormProps {
  onAddShrelloItem: ItemCreationCallback;
  onUpdateShrelloItem: ItemOperationCallback;
  onCloseForm: () => void;
  supportDepartments: ISupportDepartmentItem[];
  supportTopics: ISupportTopicItem[];
  showForm?: boolean;
  item?: IShrelloItem;
  context?: IWebPartContext;
}

interface IShrelloFormState {
  showForm: boolean;
  showLabels: boolean;
  showMembers: boolean;
  supportDepartment: string;
  supportTopic: string;
  // shrelloCategories: string[];
  supportTopics: ISupportTopicItem[];
  item: IShrelloItem;
  isNew: boolean;
  files?: File[];
}

export { IShrelloFormProps, IShrelloFormState };