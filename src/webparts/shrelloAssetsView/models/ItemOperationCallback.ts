import IShrelloItem from "./IShrelloItem";

type ItemOperationCallback = (item: IShrelloItem, attachments?: File[]) => Promise<void>;

export default ItemOperationCallback;