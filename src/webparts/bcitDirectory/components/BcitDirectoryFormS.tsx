import * as React from "react";
import styles from "./BcitDirectory.module.scss";
import { IBcitDirectoryProps } from "./IBcitDirectoryProps";
import {SPService}from '../../shared/service/bcitDirectoryServices'
import { FC } from "react";
import { sp } from "@pnp/sp/presets/all";
import { Formik, FormikProps } from "formik";
import * as yup from "yup";
import {
  IIconProps,
  PrimaryButton,
  Stack,
  TextField,
} from "office-ui-fabric-react";
import { IMyFormValuesS } from "../../shared/service/IMyFormValueS";
import { Dialog } from "@microsoft/sp-dialog";
import {ListColS} from "../../shared/service/ListColS"
import { MSGraphClient } from "@microsoft/sp-http";
import {IADirectory} from "../../shared/service/constants/IADirectory"




const BcitDirectoryFormS: FC<IBcitDirectoryProps> = (props) => {


  const cancelIcon: IIconProps = { iconName: "Cancel" };
  const addIcon: IIconProps = { iconName: "Add" };
  const CaretSolidDown: IIconProps = { iconName: "CaretSolidDown" };
  const CaretSolidUp: IIconProps = { iconName: "CaretSolidUp" };
  const _services = new SPService(props.siteUrl);
  const [showMore,setShowMore]=React.useState(false)

  sp.setup({
    spfxContext: props.context,
  });

  /** set field value and error message for all the fields */
  const getFieldProps = (formik: FormikProps<any>, field: string) => {
    return {
      ...formik.getFieldProps(field),
      errorMessage: formik.errors[field] as string,
    };
  };
 
  const addUser = (itemUser:IADirectory) => {
    props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client.api("/users").version("v1.0").post(itemUser);
      });
    Dialog.alert(" you add new user check your list of active user :)");
    console.log(itemUser,'iradddddddddddddddd')
    
  };
   
  const createRecord = (record: any, helpers: any) => {
    let item: ListColS = {
      FIRST_NAME: record.firstName,
      LAST_NAME: record.lastName,
      DISPLAY_NAME: record.firstName + " " + record.lastName,
      EMAIL: record.email,
      SEND_INVITATION_OPTION: record.sendInvitationOption,
      REDIRECT_URL: record.redirectUrl,
      JOB_TITLE: record.jobTitle,
      DEPARTEMENT: record.department,
      MANAGER: record.manager,
    };

    const itemUser: IADirectory = {
      accountEnabled: true,
      displayName: record.firstName + "." + record.lastName,
      mailNickname: record.firstName,
      userPrincipalName: `${record.firstName}.${record.lastName}@bestconsultingit.onmicrosoft.com`,
      passwordProfile: {
        forceChangePasswordNextSignIn: true,
        password: "Best@test2021",
      },
    };

    addUser(itemUser);

    _services
      .createTask("Tasks", item)
      .then(() => {
        Dialog.alert(
          "data added in list sharepoint successfully check your list << Users >>!!! :)"
        );
        helpers.resetForm();
      })
      .catch((err) => {
        console.error(err);
        throw err;
      });
  };

  const validate = yup.object().shape({
    firstName: yup.string().required("This information is required"),
    lastName: yup.string().required("This information is required"),
    email: yup.string().email("Invalid email").required("Required"),
  });

  const initialValues: IMyFormValuesS = {
    firstName: "",
    lastName: "",
    displayName: "",
    email: "",
    sendInvitationOption: "",
    redirectUrl: "",
    jobTitle: "",
    department: "",
    manager: "",
  };
  return (
    <Formik
      initialValues={initialValues}
      validationSchema={validate}
      onSubmit={createRecord}
    >
      {(formik) => (
        <div className={styles.reactFormik}>
          <Stack>
            

            <Stack horizontal horizontalAlign="space-evenly">
              <TextField
                name="firstName"
                label="First Name"
                id="firstName"
                {...getFieldProps(formik, "firstName")}
              />
              <TextField
                name="lastName"
                label="Last Name"
                id="lastName"
                {...getFieldProps(formik, "lastName")}
              />
              <TextField
                name="displayName"
                label="Display Name"
                id="displayName"
                value={formik.values.firstName + " " + formik.values.lastName}
               // {...getFieldProps(formik, "displayName")}
                disabled
              />
            </Stack>

            <Stack horizontal horizontalAlign="space-evenly">
              <TextField
                name="email"
                label="Email"
                {...getFieldProps(formik, "email")}
              />
              <TextField
                name="sendInvitationOption"
                label="Send Invitation"
                {...getFieldProps(formik, "sendInvitationOption")}
              />
              <TextField
                name="redirectUrl"
                label="Redirect URL"
                {...getFieldProps(formik, "redirectUrl")}
              />
            </Stack>
               {showMore ?  
               <Stack>
                <Stack horizontal horizontalAlign="center">
               <PrimaryButton
               id="buttonMoreInformations"
               type="button"
               name="buttonMoreInformations"
               text="Less Informations"
               iconProps={CaretSolidDown}
               className={styles.btnsForm}
               onClick={()=>setShowMore(!showMore)}
             />
             </Stack>
              <Stack
              horizontal
              horizontalAlign="space-evenly"
              id="moreInformationsBloc"
              hidden={true}
            >
              <TextField
                name="jobTitle"
                label="Job Title"
                {...getFieldProps(formik, "jobTitle")}
              />
              <TextField
                name="department"
                label="Department"
                {...getFieldProps(formik, "department")}
              />
              <TextField
                name="manager"
                label="Manager"
                {...getFieldProps(formik, "manager")}
              />
            </Stack>
            </Stack>
            :
            <Stack horizontal horizontalAlign="center">
              <PrimaryButton
                id="buttonMoreInformations"
                type="button"
                name="buttonMoreInformations"
                text="More Informations"
                iconProps={CaretSolidDown}
                className={styles.btnsForm}
                onClick={()=>setShowMore(!showMore)}
              />
            </Stack>
            }
          </Stack>
          <Stack horizontal horizontalAlign="center">
            <PrimaryButton
              type="submit"
              text="Save"
              iconProps={addIcon}
              className={styles.btnsForm}
              onClick={formik.handleSubmit as any}
            />
            <PrimaryButton
              text="Cancel"
              iconProps={cancelIcon}
              className={styles.btnsForm}
              onClick={formik.handleReset}
            />
          </Stack>
        </div>
      )}
    </Formik>
  );
};
export default BcitDirectoryFormS
