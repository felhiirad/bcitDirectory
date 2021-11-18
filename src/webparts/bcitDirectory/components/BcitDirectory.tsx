import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBcitDirectoryProps } from "./IBcitDirectoryProps";
import BcitDirectoryForm from "./BcitDirectoryForm";
import { MSGraphClient } from "@microsoft/sp-http";
import { useEffect, useState } from "react";
import { IUserItem, IUserData } from "./IBcitDirectoryState";
import ListData from "./ListData";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { Dialog } from "@microsoft/sp-dialog";






// const controlClass = mergeStyleSets({
//   control: {
//     margin: "0 0 15px 0",
//     maxWidth: "300px",
//   },
// });
const ListItemsWebPartContext = React.createContext<WebPartContext>(null);

const BcitDirectory: React.FC<IBcitDirectoryProps> = (props) => {
  const [user, setUser] = useState<IUserItem>();
  const [data, setData] = useState<IUserData>();

  const users = {
    accountEnabled: true,
    displayName: "test",
    mailNickname: "amine",
    userPrincipalName: "test.amine@bestconsultingit.onmicrosoft.com",
    passwordProfile: {
      forceChangePasswordNextSignIn: true,
      password: "Iradfelhi485",
    },
  };
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

//function to add user
   const addUsers = () => {
    props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("/users")
          .version("v1.0")
          .post(users);
      });
      Dialog.alert(
        " you add new user check your list of active user :)"
      );
  };

  return (
    <>
      <BcitDirectoryForm
        context={props.context}
        description={props.description}
        siteUrl={props.siteUrl}
      />
      <div> submited by {user && user.displayName}</div>
      <div> mail :{user && user.mail}</div>

      <ListData
        context={props.context}
        description={props.description}
        siteUrl={props.siteUrl}
      />
      <PrimaryButton text="Add user" onClick={() =>addUsers() }></PrimaryButton>
    </>
  );
};

export default BcitDirectory;
