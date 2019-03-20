import { PrincipalInfo } from "@pnp/sp";

interface ISupportTopicItem {
  Id?: number;
  Title?: string;
  TASCAbbreviation?: string;
  SupportTopicDescription?: string;
  Language?: string;
  Classification?: string;
  Vendor?: string;
  VendorContact?: string;
  VendorEmail?: string;
  VendorWebsite?: string;
  OnRequestForm?: boolean;
  ResponsibleSupervisorId?: number;
  ResponsibleSupervisor?: PrincipalInfo;
  PrimarySupportId?: number;
  PrimarySupport?: PrincipalInfo;
  SecondarySupportId?: number;
  SecondarySupport?: PrincipalInfo;
  SupportContactId?: number;
  SupportContact?: PrincipalInfo;
  SupportDepartmentId?: number;
  AccessContact?: number;
  TicketCount?: number;
}

export default ISupportTopicItem;