import IShrelloItem from "./IShrelloItem";

type ItemCreationCallback = (item: IShrelloItem, attachments?: File[]) => Promise<void>;

export default ItemCreationCallback;