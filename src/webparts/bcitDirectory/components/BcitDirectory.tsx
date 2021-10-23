import * as React from 'react';
import styles from './BcitDirectory.module.scss';
import { IBcitDirectoryProps } from './IBcitDirectoryProps';
import { IBcitDirectoryState } from './IBcitDirectoryState';
import { SPService } from '../../shared/service/SPService';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Stack } from 'office-ui-fabric-react/lib/Stack';
import { Formik, FormikProps, Field } from 'formik';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { DatePicker,  mergeStyleSets, PrimaryButton, IIconProps } from 'office-ui-fabric-react';
import { sp } from '@pnp/sp/presets/all';
import { Dialog } from '@microsoft/sp-dialog';


const controlClass = mergeStyleSets({
  control: {
    margin: '0 0 15px 0',
    maxWidth: '300px',
  },
});
interface MyFormValues {
  createdBy: string,
  updatedBy: string,
  errorMsg: string,
  picked: string,
  startDate: any,
  endDate: any
   }
export default class BcitDirectory extends React.Component<IBcitDirectoryProps, IBcitDirectoryState, {}> {
  private cancelIcon: IIconProps = { iconName: 'Cancel' };
  private saveIcon: IIconProps = { iconName: 'Save' };
  private _services: SPService = null;
  constructor(props: Readonly<IBcitDirectoryProps>) {
    super(props);
    this.state = {
      startDate: null,
      endDate: null,
    };
    sp.setup({
      spfxContext: this.props.context
    });
    this._services = new SPService(this.props.siteUrl);
  }
  /** set field value and error message for all the fields */
  private getFieldProps = (formik: FormikProps<any>, field: string) => {
    return { ...formik.getFieldProps(field), errorMessage: formik.errors[field] as string }
  }
  /** create record in sharepoint list */

  public async createRecord(record: any) {
      await this._services.createTask("Tasks", {
          CREATED_BY: record.createdBy,
          UPDATED_BY: record.updatedBy,
          SUCCESS:record.picked,
          CREATION_DATE: record.startDate,
          UPDATE_DATE: new Date(record.endDate),
          ERROR_MSG: record.errorMsg,
    }).then((data) => {
      Dialog.alert("Added to the list Sharepoint 'UsersLog' successfully !!!");
      return data;

    }).catch((err) => {
      console.error(err);
      throw err;
    });
  }
  public render(): React.ReactElement<IBcitDirectoryProps> {
    const initialValues :MyFormValues={
          createdBy: '',
          updatedBy: '',
          errorMsg: '',
          picked: '',
          startDate: null,
          endDate: null
    }
    return (
      < div className="resetForm">
        <Formik initialValues={initialValues}
          //validationSchema={validate}
          onSubmit={this.createRecord}>
          {formik => (
            <div className={styles.reactFormik}>
              <Stack>
                <Label className={styles.lblForm}>Best Consulting IT Directory</Label>
              
                <Label className={styles.lblForm}>Created By</Label>
              <TextField name = "createdBy" placeholder="Created BY"
                 />
               
                <Label className={styles.lblForm}>Created Date</Label>
              <DatePicker
                className={controlClass.control}
                id="startDate"
                value={formik.values.startDate}
                onSelectDate={(date) => formik.setFieldValue('startDate', date)}
              />
                
                <Label className={styles.lblForm}>Updated By</Label>
              <TextField
                 name="updatedBy"
                 placeholder="Updated By" 
                />
                
              <Label className={styles.lblForm}>Updated Date</Label>
              <DatePicker
                className={controlClass.control}
                id="endDate"
                value={formik.values.startDate}
                onSelectDate={(date) => formik.setFieldValue('endDate', date)}
              />
                
                <Label className={styles.lblForm}>Success</Label>
                <div   >
                  <label >
                    <Field type="radio" name="picked" value="Yes"
                     chacked
                    />
                    Yes
                  </label>
                  <label >
                    <Field type="radio" name="picked" value="No"
                      chacked
                    />
                    No
                  </label>
                </div>
                <Label className={styles.lblForm}>Error Msg</Label>
              <TextField
                multiline
                rows={6}
                name="errorMsg"
                />
              </Stack>
              <PrimaryButton
                 type="submit"
                 text="Save"
                 iconProps={this.saveIcon}
                 className={styles.btnsForm}
              />
              <PrimaryButton
                 text="Cancel"
                 iconProps={this.cancelIcon}
                 className={styles.btnsForm}
                 onClick={formik.handleReset} 
              />
            </div>
          )
          }
        </Formik >
      </div >
    );
  }
}
