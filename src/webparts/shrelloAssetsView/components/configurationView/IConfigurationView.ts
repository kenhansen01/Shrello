interface IConfigurationViewProps {
  icon?: string;
  iconText?: string;
  description?: string;
  buttonLabel?: string;
  onConfigure?: () => void;
}

interface IConfigurationViewState {}

export { IConfigurationViewProps, IConfigurationViewState };