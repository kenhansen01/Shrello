import { SPRest, Web } from "@pnp/sp";

export default class ShrelloAssetsViewConstants {
  private _sp: SPRest;
  private _web: Web;

  public set sp(value: SPRest) {
    this._sp = value;
    this._web = value.web;
  }

  public get sp(): SPRest {
    return this._sp;
  }

  public set web(value: Web) {
    this._web = value;
  }

  public get web(): Web {
    return this._web;
  }
}