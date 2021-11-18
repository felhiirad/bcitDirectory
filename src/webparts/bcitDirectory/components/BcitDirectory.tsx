import * as React from "react";
import { IBcitDirectoryProps } from "./IBcitDirectoryProps";
import BcitDirectoryForm from "./BcitDirectoryForm";
import { MSGraphClient } from "@microsoft/sp-http";
import { useEffect, useState } from "react";
import { IUserItem, IUserData } from "./IBcitDirectoryState";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { Dialog } from "@microsoft/sp-dialog";
import BcitDirectoryFormS from "./BcitDirectoryFormS";
import { Dropdown, Icon, Label, Stack } from "office-ui-fabric-react";
import styles from "./BcitDirectory.module.scss";

// const controlClass = mergeStyleSets({
//   control: {
//     margin: "0 0 15px 0",
//     maxWidth: "300px",
//   },
// });
//const ListItemsWebPartContext = React.createContext<WebPartContext>(null);

const BcitDirectory: React.FC<IBcitDirectoryProps> = (props) => {
  const [user, setUser] = useState<IUserItem>();
  const [data, setData] = useState<IUserData>();
  const [optionValue, setOptionValue] = useState("");
  const Info = () => <Icon iconName="Info" />;

  
  // function
  useEffect(() => {
    props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("me")
          .version("v1.0")
          .select("displayName,mail,companyName")
          .get((err, res) => {
            if (err) {
              console.error(" MSGraph API Error");
              console.error(err);
              return;
            }
            setUser({
              displayName: res.displayName,
              mail: res.mail,
              companyName: res.companyName,
            });
            console.log("response", res);
          });
      });
  }, []);

 
 
  console.log(optionValue, "Account typeeeeeeeeeeeee");

  return (
    <>
      <Label className={styles.lblForm}>
        Best Consulting IT Directory <Info />
      </Label>
      <Stack horizontal horizontalAlign="space-evenly">
        <Dropdown
          id="addType"
          label="Account type :"
          onChange={(e, selectedOption) => {
            setOptionValue(selectedOption.text)
          }}
          options={[
            { key: "guest", text: "Guest", isSelected: true },
            { key: "microsoft365", text: "Microsoft 365" },
          ]}
        />
      </Stack>

      {optionValue == "Microsoft 365" ? (
        <BcitDirectoryFormS
          context={props.context}
          description={props.description}
          siteUrl={props.siteUrl}
        />
      ) : (
        <BcitDirectoryForm
          context={props.context}
          description={props.description}
          siteUrl={props.siteUrl}
        />
      )}
      <br></br>
    </>
  );
};
export default BcitDirectory;
