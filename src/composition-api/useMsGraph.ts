import { injectStrict } from '@/utils/typescript-utils';
import { AxiosKey } from '@/types/symbols';

const baseUrl = 'https://graph.microsoft.com/v1.0/me';

export function useMsGraph() {
  const axios = injectStrict(AxiosKey);

  async function getGraphProfile() {
    try {
      const { data } = await axios.get<microsoftgraph.User>(baseUrl);
      return data;
    } catch (error) {
      console.log('Error getting user profile: ', error);
    }
  }

  async function getGraphDriveItems() {
    try {
      const { data } = await axios.get(`${baseUrl}/drive/root/children`);
      return data;
    } catch (error) {
      console.log('Error getting drive items: ', error);
    }
  }

  async function getGraphExcel(id: string) {
    try {
      const { data } = await axios.get(
        `${baseUrl}/drive/items/${id}/workbook/worksheets`,
      );
      return data;
    } catch (error) {
      console.log('Error getting worksheets: ', error);
    }
  }

  async function postGraphExcelRow() {
    try {
      const payload = {
        values: [['alex darrow', '123', 'adarrow@tenant.onmicrosoft.com']],
      };

      const { data } = await axios.post(
        `${baseUrl}/drive/items/01K6AWNMKR4C6YCVVJHZG3ZTWJVJDXQCN6/workbook/worksheets/Sheet1/tables/Table1/rows`,
        payload,
      );

      return data;
    } catch (error) {
      console.log('Error creating row: ', error);
    }
  }

  return {
    getGraphProfile,
    getGraphDriveItems,
    getGraphExcel,
    postGraphExcelRow,
  };
}
