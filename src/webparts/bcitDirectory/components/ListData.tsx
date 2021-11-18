import * as React from 'react';
import {IBcitDirectoryProps} from "./IBcitDirectoryProps";
import {escape} from "@microsoft/sp-lodash-subset";
import {PrimaryButton} from "office-ui-fabric-react/lib/Button";
import {MSGraphClient} from "@microsoft/sp-http";
import {DetailsList,ThemeSettingName} from "office-ui-fabric-react";
import {IUserItem,IUserData} from "./IBcitDirectoryState" 
import { useEffect,useState } from "react";
import { render } from 'react-dom';

export interface IList{
    createdBy: string,
    updatedBy: string,
    errorMsg: string,
    creationDate:string ,
    updateDate:string,
    success:string,
}
export interface IUserState{
    listDataState:IList[];
    visible:boolean;
}
export default class ListData extends React.Component<
IBcitDirectoryProps,
IUserState,{}
>
{
    constructor(props:IBcitDirectoryProps){
        super(props);
        this.state={listDataState:[],visible:true}
    }

public listData:IList[]=[];
public GetListData=():void=>{
         this.props.context.msGraphClientFactory
         .getClient()
         .then((client:MSGraphClient)=>{
            client
            .api("https://graph.microsoft.com/v1.0/sites/bestconsultingit.sharepoint.com,00910f08-a052-4034-8ca3-a3e0239de48d,6fe45d08-dabc-4b96-8ee8-d1cf69fb7a2d/lists/%7B37bc2bf6-7f08-464f-b4bd-99a3cbfbed91%7D/items")
            .expand(`fields($select=CREATED_BY,CREATION_DATE,UPDATED_BY,UPDATE_DATE,SUCCESS,ERROR_MSG)`)
            .version("v1.0")
            //.select("CREATED_BY,UPDATED_BY,CREATION_DATE,UPDATED_DATE,SUCCESS,ERROR_MSG")
             .get((err,res)=>{
                if(err){
                   console.error(" MSGraph API Error to get data from list sharepoint" )
                console.error(err);
                   
                  }
                  console.log("Response list data",res)
                  res.value.map((result)=>{
                      this.listData.push({
                        createdBy:result.fields.CREATED_BY ,
                        updatedBy:result.fields.UPDATED_BY ,
                        errorMsg:result.fields.ERROR_MSG ,
                        creationDate:result.fields.CREATION_DATE ,
                        updateDate:result.fields.UPDATE_DATE,
                        success:result.fields.SUCCESS,
                      }) 
                  })
                 this.setState({listDataState:this.listData})   
             })
         })
      }
      public render():React.ReactElement<IBcitDirectoryProps>{
        let slider= this.state.visible ? (
            <DetailsList items={this.state.listDataState}></DetailsList>
            
            ):(
            null
            && this.setState({visible:true})
            );
        //  const buttonText=this.state.visible? "Show Data From UserLog " :"Hide Data";

          return (
              <div>
              <PrimaryButton text="Get Data From UserLog" onClick={this.GetListData}>
              </PrimaryButton>
                     {slider}
             <PrimaryButton text="Hide list" onClick={()=>this.setState({visible:false})}></PrimaryButton>
              </div>
          )
      }
    }

    
// Application (client) ID
// 6e0a9cfb-eee9-4247-a704-19a74a3adb35
// Directory (tenant) ID
// fecbf840-99cf-4d6c-a783-eb40381b3138

