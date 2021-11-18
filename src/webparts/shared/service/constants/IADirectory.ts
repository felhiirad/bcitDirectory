export  interface IADirectory
{
    accountEnabled: boolean;
    displayName: string;
    mailNickname: string;
    userPrincipalName: string;
    passwordProfile: {
      forceChangePasswordNextSignIn: boolean;
      password: string
  }
}