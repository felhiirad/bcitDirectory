

export interface IBcitDirectoryState {
    startDate?: any;
    endDate?: any;

  }
  export interface IUserItem {
    displayName:string;
    mail:string;
    companyName:string;
   
  }
  
  export interface IUserData {
    CREATED_BY:string ;
    UPDATED_BY:string ;
    CREATION_DATE:Date;
    UPDATED_DATE:Date;
    SUCCESS:string;
    ERROR_MSG:string ;
  }