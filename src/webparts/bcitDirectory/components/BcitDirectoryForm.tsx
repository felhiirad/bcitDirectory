import * as React from "react";
import * as ReactDom from 'react-dom';
import { useState, FC } from "react";
import styles from "./BcitDirectory.module.scss";
import { IBcitDirectoryProps } from "./IBcitDirectoryProps";
import { SPService } from "../../shared/service/SPService";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Stack } from "office-ui-fabric-react/lib/Stack";
import { Formik, FormikProps, Field } from "formik";
import { Label } from "office-ui-fabric-react/lib/Label";
import {
  DatePicker,
  mergeStyleSets,
  PrimaryButton,
  IIconProps,
} from "office-ui-fabric-react";
import { sp } from "@pnp/sp/presets/all";
import { Dialog } from "@microsoft/sp-dialog";
import { MyFormValues } from "../../shared/service/MyFormValues";
import { ListCol } from "../../shared/service/ListCol";
import * as yup from "yup";
import {WebPartContext} from '@microsoft/sp-webpart-base';
import Recaptcha from "react-recaptcha";






const controlClass = mergeStyleSets({
    control: {
      margin: "0 0 15px 0",
      maxWidth: "300px",
    },
  });
  
  export const  ListItemsWebPartContext=React.createContext<WebPartContext>(null);
  
  const BcitDirectoryForm: React.FC<IBcitDirectoryProps> = (props) => {
    const cancelIcon: IIconProps = { iconName: "Cancel" };
    const saveIcon: IIconProps = { iconName: "Save" };
    const _services = new SPService(props.siteUrl);
  
    const [startDate, setStartDate] = useState(null);
    const [endDate, setEndDate] = useState(null);
  
    sp.setup({
        spfxContext: props.context
    });
  
    //const  _services = new SPService(props.siteUrl);
    //let  createRecord = createRecord.bind(this);
  
    /** set field value and error message for all the fields */
    const getFieldProps = (formik: FormikProps<any>, field: string) => {
      return {
        ...formik.getFieldProps(field),
        errorMessage: formik.errors[field] as string,
      };
    };
    /** create record in sharepoint list */
  
    async function createRecord(record: any) {
      const item: ListCol = {
        CREATED_BY: record.createdBy,
        UPDATED_BY: record.updatedBy,
        SUCCESS: record.picked,
        CREATION_DATE: record.startDate,
        UPDATE_DATE: new Date(record.endDate),
        ERROR_MSG: record.errorMsg,
      };
  
      await _services
        .createTask("Tasks", item)
        .then((data) => {
          Dialog.alert(
            "data added in list sharepoint >>successfully  check your list !!!  :)"
          );
          return data;
        })
        .catch((err) => {
          console.error(err);
          throw err;
        });
    }
  
    /** validations */
    const validate = yup.object().shape({
      createdBy: yup.string().required("created by is required"),
      errorMsg: yup
        .string()
        .min(15, "Minimum required 15 characters")
        .required("error msg is required"),
      updatedBy: yup.string().required("updated by  is required"),
      startDate: yup
        .date()
        .required("Please select the  date of creation")
        .nullable(),
      endDate: yup
        .date()
        .required("Please select the date of updating")
        .nullable(),
      picked: yup.string().required("please success field is required "),
      recaptcha: yup.string().required("reCAPTCHA is required "),

    });
    const initialValues: MyFormValues = {
      createdBy: "",
      updatedBy: "",
      errorMsg: "",
      picked: "",
      startDate: null,
      endDate: null,
      recaptcha:""
    };
    React.useEffect(()=>{
      const script=document.createElement("script")
      script.src="https://www.google.com/recaptcha/api.js"
      script.async = true;
      script.defer = true;
    document.body.appendChild(script);
    })
    return (
     
      <Formik
        initialValues={initialValues}
        validationSchema={validate}
        //onSubmit={this.createRecord}
        onSubmit={(values, helpers) => {
          console.log("SUCCESS!! :-)\n\n" + JSON.stringify(values, null, 4));
          createRecord(values).then((response) => {
            helpers.resetForm();
          });
        }}
      >
        {(formik) => (
          <div className={styles.reactFormik}>
            <Stack>

              <Label className={styles.lblForm}>
                Best Consulting IT Directory
              </Label>
              
  
              <Label className={styles.lblForm}>Created By</Label>
              <TextField
                name="createdBy"
                {...getFieldProps(formik, "createdBy")}
              />
              
  
              <Label className={styles.lblForm}>Created Date</Label>
              <DatePicker
                className={controlClass.control}
                id="startDate"
                value={formik.values.startDate}
                textField={{ ...getFieldProps(formik, "startDate") }}
                onSelectDate={(date) => formik.setFieldValue("startDate", date)}
              />
          
  
              <Label className={styles.lblForm}>Updated By</Label>
              <TextField
                name="updatedBy"
                {...getFieldProps(formik, "updatedBy")}
              />
  
              <Label className={styles.lblForm}>Updated Date</Label>
              <DatePicker
                className={controlClass.control}
                id="endDate"
                value={formik.values.startDate}
                textField={{ ...getFieldProps(formik, "endDate") }}
                onSelectDate={(date) => formik.setFieldValue("endDate", date)}
              />
  
              <Label className={styles.lblForm}>Success</Label>
  
              <label>
                <Field type="radio" name="picked" value="Yes" chacked />
                Yes
              </label>
              <label>
                <Field type="radio" name="picked" value="No" />
                No
              </label>
  
              <Label className={styles.lblForm}>Error Msg</Label>
              <TextField
                multiline
                rows={3}
                name="errorMsg"
                {...getFieldProps(formik, "errorMsg")}
              />
               <Recaptcha
                  sitekey="6Lfe3fscAAAAAEgIAzMIUm7K0a0s7WxDVDCIAsT_"
                  render="explicit"
                  theme="dark"
                  verifyCallback={(response) => { formik.setFieldValue("recaptcha", response); }}
                  onloadCallback={() => { console.log("done loading!"); }}
                />
            </Stack>
           
                 
            <PrimaryButton
              type="submit"
              text="Save"
              iconProps={saveIcon}
              className={styles.btnsForm}
              onClick={formik.handleSubmit as any}
            />
            <PrimaryButton
              text="Cancel"
              iconProps={cancelIcon}
              className={styles.btnsForm}
              onClick={formik.handleReset}
            />
          </div>
        )}
      </Formik>   


    );
  };

  export default BcitDirectoryForm;