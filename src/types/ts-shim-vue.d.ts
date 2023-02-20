import { ComponentCustomProperties } from 'vue';
import type {
  InteractionStatus,
  PublicClientApplication,
  AccountInfo,
} from '@azure/msal-browser';

interface IMSAL {
  instance: PublicClientApplication;
  inProgress: InteractionStatus.Startup;
  accounts: AccountInfo[];
}

declare module '@vue/runtime-core' {
  interface ComponentCustomProperties {
    $msal: IMSAL;
  }
}
