import * as React from "react";
import { IDataContext } from "./IDataContext";

export const DataContext: any = React.createContext<any>(null);
export const ContextProvider = DataContext.Provider;
export const ContextConsumer = DataContext.Consumer;