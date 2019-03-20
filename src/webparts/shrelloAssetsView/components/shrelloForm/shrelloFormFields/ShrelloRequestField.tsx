import * as React from "react";
import { stateToHTML } from "draft-js-export-html";
import {
  EditorState,
  convertFromHTML,
  ContentState,
  Editor,
  RichUtils,
  getDefaultKeyBinding,
  DraftHandleValue,
  DraftBlockType,
  DraftInlineStyleType,
  DraftInlineStyle,
  SelectionState,
  ContentBlock
} from "draft-js";

import { IShrelloItem } from "../../../models";
import "draft-js/dist/Draft.css";
import "./ShrelloRequestField.css";

// custom overrides for "code" style.
const styleMap: any = {
  CODE: {
    backgroundColor: "rgba(0, 0, 0, 0.05)",
    fontFamily: "\"Inconsolata\", \"Menlo\", \"Consolas\", monospace",
    fontSize: 16,
    padding: 2,
  },
};

const BLOCK_TYPES: { label: string, style: DraftBlockType }[] = [
  {label: "H1", style: "header-one"},
  {label: "H2", style: "header-two"},
  {label: "H3", style: "header-three"},
  // {label: "H4", style: "header-four"},
  // {label: "H5", style: "header-five"},
  // {label: "H6", style: "header-six"},
  {label: "Blockquote", style: "blockquote"},
  {label: "UL", style: "unordered-list-item"},
  {label: "OL", style: "ordered-list-item"},
  {label: "Code Block", style: "code-block"},
];

const INLINE_STYLES: { label: string, style: DraftInlineStyleType }[] = [
  {label: "Bold", style: "BOLD"},
  {label: "Italic", style: "ITALIC"},
  {label: "Underline", style: "UNDERLINE"},
  // {label: "Monospace", style: "CODE"},
];

// const InlineStyleControls: (props: any) => JSX.Element = (props): JSX.Element => {
//   const editorState: EditorState = props.editorState;
//   const currentStyle: DraftInlineStyle = editorState.getCurrentInlineStyle();
//   return (
//     <div className="RichEditor-controls">
//       {INLINE_STYLES.map((type) =>
//         <StyleButton
//           key={type.label}
//           active={currentStyle.has(type.style)}
//           label={type.label}
//           onToggle={props.onToggle}
//           style={type.style}
//         />
//       )}
//     </div>
//   );
// };

const BlockStyleControls: (props: any) => JSX.Element = (props): JSX.Element => {
  const editorState: EditorState = props.editorState;
  const selection: SelectionState = editorState.getSelection();
  const currentStyle: DraftInlineStyle = editorState.getCurrentInlineStyle();
  const blockType: DraftBlockType = editorState
    .getCurrentContent()
    .getBlockForKey(selection.getStartKey())
    .getType();
  return (
    <div className="RichEditor-controls">
      {INLINE_STYLES.map((type) =>
        <StyleButton
          key={type.label}
          active={currentStyle.has(type.style)}
          label={type.label}
          onToggle={props.onStyleToggle}
          style={type.style}
        />
      )}
      {BLOCK_TYPES.map((type) =>
        <StyleButton
          key={type.label}
          active={type.style === blockType}
          label={type.label}
          onToggle={props.onBlockToggle}
          style={type.style}
        />
      )}
    </div>
  );
};

export interface IShrelloRequestProps {
  item?: IShrelloItem;
  onSave?: (rqHtml: string) => void;
}

export default class ShrelloRequestField extends React.Component<IShrelloRequestProps, {editorState: EditorState}> {
  private _initial: EditorState = EditorState.createEmpty();
  private _editorRef: Editor;

  constructor(props: IShrelloRequestProps) {
    super(props);

    const blocksFromHTML: {contentBlocks: Array<ContentBlock>, entityMap: any } = props.item.Body
      ? convertFromHTML(props.item.Body)
      : null;
    const rqState: ContentState = blocksFromHTML
      ? ContentState.createFromBlockArray(
        blocksFromHTML.contentBlocks,
        blocksFromHTML.entityMap
      )
      : null;

    this._initial = rqState
      ? EditorState.createWithContent(rqState)
      : this._initial;
    this.state = {
      editorState: this._initial || null
    };
    this._focus = this._focus.bind(this);
    this._onChange = this._onChange.bind(this);
    this._handleKeyCommand = this._handleKeyCommand.bind(this);
    this._mapKeyToEditorCommand = this._mapKeyToEditorCommand.bind(this);
    this._toggleBlockType = this._toggleBlockType.bind(this);
    this._toggleInlineStyle = this._toggleInlineStyle.bind(this);
  }

  private _focus = () => this._editorRef.focus();

  private _onChange = (editorState: EditorState) => {
    const rqHtml: string = stateToHTML(editorState.getCurrentContent());
    this.props.onSave(rqHtml);
    this.setState({editorState});
  }

  private _handleKeyCommand = (command, editorState: EditorState): DraftHandleValue => {
    const newState: EditorState = RichUtils.handleKeyCommand(editorState, command);
    if (newState) {
      this._onChange(newState);
      return "handled";
    }
    return "not-handled";
  }
  private _mapKeyToEditorCommand = (e) => {
    if (e.keyCode === 9 /* TAB */) {
      const newEditorState: EditorState = RichUtils.onTab(
        e,
        this.state.editorState,
        4, /* maxDepth */
      );
      if (newEditorState !== this.state.editorState) {
        this._onChange(newEditorState);
      }
      return;
    }
    return getDefaultKeyBinding(e);
  }
  private _toggleBlockType = (blockType) => {
    this._onChange(
      RichUtils.toggleBlockType(
        this.state.editorState,
        blockType
      )
    );
  }
  private _toggleInlineStyle = (inlineStyle) => {
    this._onChange(
      RichUtils.toggleInlineStyle(
        this.state.editorState,
        inlineStyle
      )
    );
  }
  public render(): JSX.Element {
    const editorState: EditorState = this.state.editorState;
    // if the user changes block type before entering any text, we can
    // either style the placeholder or hide it. Let's just hide it now.
    let className: string = "RichEditor-editor";
    var contentState: ContentState = editorState.getCurrentContent();
    if (!contentState.hasText()) {
      if (contentState.getBlockMap().first().getType() !== "unstyled") {
        className += " RichEditor-hidePlaceholder";
      }
    }
    return (
      <div className="RichEditor-root">
        <BlockStyleControls
          editorState={editorState}
          onBlockToggle={this._toggleBlockType}
          onStyleToggle={this._toggleInlineStyle}
        />
        <div className={className} onClick={this._focus}>
          <Editor
            blockStyleFn={getBlockStyle}
            customStyleMap={styleMap}
            editorState={editorState}
            handleKeyCommand={this._handleKeyCommand}
            keyBindingFn={this._mapKeyToEditorCommand}
            onChange={this._onChange}
            placeholder="Tell a story..."
            ref={el => this._editorRef = el}
            spellCheck={true}
          />
        </div>
      </div>
    );
  }
}

function getBlockStyle(block: any): string {
  switch (block.getType()) {
    case "blockquote": return "RichEditor-blockquote";
    default: return null;
  }
}

export interface IStyleButtonProps {
  onToggle: (style: any) => void;
  style: any;
  active: boolean;
  label: string;
}
class StyleButton extends React.Component<IStyleButtonProps, {}> {
  constructor() {
    super();
    this._onToggle = this._onToggle.bind(this);
  }
  private _onToggle = (e) => {
    e.preventDefault();
    this.props.onToggle(this.props.style);
  }
  public render(): JSX.Element {
    let className: string = "RichEditor-styleButton";
    if (this.props.active) {
      className += " RichEditor-activeButton";
    }
    return (
      <span className={className} onMouseDown={this._onToggle}>
        {this.props.label}
      </span>
    );
  }
}