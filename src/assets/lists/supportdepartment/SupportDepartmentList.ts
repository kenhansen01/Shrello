import SupportDepartmentListFields from "./SupportDepartmentListFields";
import { IListDef } from "../../../utilities/ListUtility";

const SupportDepartmentList: IListDef = {
  title: "Support Departments",
  description: "Support Departments that use TASC.",
  template: 100,
  enableContentTypes: true,
  additionalSettings: {
    EnableVersioning: true
  },
  fieldDefs: SupportDepartmentListFields,
  defaultViewName: "All Items"
};

export default SupportDepartmentList;
