import * as React from "react";
import {WebPartContext} from '@microsoft/sp-webpart-base';
import {IBcitDirectoryProps} from "./IBcitDirectoryProps";
import BcitDirectoryForm from "./BcitDirectoryForm";


// const controlClass = mergeStyleSets({
//   control: {
//     margin: "0 0 15px 0",
//     maxWidth: "300px",
//   },
// });
export const  ListItemsWebPartContext=React.createContext<WebPartContext>(null);

const BcitDirectory: React.FC<IBcitDirectoryProps> = (props) => {
  
  return (
        < BcitDirectoryForm context={props.context} description={props.description} siteUrl={props.siteUrl} />
  );
};

export default BcitDirectory;
