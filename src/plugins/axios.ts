import type { App } from 'vue';
import axios from 'axios';
import { AxiosKey } from '@/types/symbols';
import { loginRequest } from '../authConfig';
import { InteractionRequiredAuthError } from '@azure/msal-browser';

const axiosInstance = axios.create();

export default {
  install: async (app: App): Promise<void> => {
    const instance = app.config.globalProperties.$msal.instance;

    axiosInstance.interceptors.request.use(
      (config) => {
        return instance
          .acquireTokenSilent({
            ...loginRequest,
          })
          .then((response) => {
            config.headers.Authorization = response.accessToken;
            return Promise.resolve(config);
          })
          .catch(async (e) => {
            console.log('acquireTokenSilent error: ', e);

            if (e instanceof InteractionRequiredAuthError) {
              await instance.acquireTokenRedirect(loginRequest);
            }
            throw e;
          });
      },
      (error) => {
        console.error(error);
        return Promise.reject(error);
      },
    );

    app.config.globalProperties.$axios = axiosInstance;
    app.provide(AxiosKey, axiosInstance);
  },
};

export { axiosInstance };
