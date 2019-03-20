import SupportTopicListFields from "./SupportTopicListFields";
import { IListDef } from "../../../utilities/ListUtility";

const SupportTopicList: IListDef = {
  title: "Support Topics",
  description: "Support Topics supported by TASC",
  template: 100,
  enableContentTypes: true,
  additionalSettings: {
    EnableVersioning: true,
    EnableAttachments: true,
    Hidden: false,
    NoCrawl: false
  },
  fieldDefs: SupportTopicListFields,
  defaultViewName: "All Items"
};

export default SupportTopicList;
