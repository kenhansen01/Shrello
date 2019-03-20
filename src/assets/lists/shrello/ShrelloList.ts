import ShrelloListFields from "./ShrelloListFields";
import { IListDef } from "../../../utilities/ListUtility";

const ShrelloList: IListDef = {
  title: "TASC Tickets",
  description: "TASC issues and requests.",
  template: 171,
  enableContentTypes: true,
  additionalSettings: {
    EnableVersioning: true,
    EnableAttachments: true,
    Hidden: false,
    NoCrawl: false,
    OnQuickLaunch: true
  },
  fieldDefs: ShrelloListFields,
  defaultViewName: "All Tasks"
};

export default ShrelloList;
