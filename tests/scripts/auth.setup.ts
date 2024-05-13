// import { test as setup } from '@playwright/test';

// const authFile = 'playwright/.auth/user.json';

// setup('authenticate', async ({ page }) => {

//   await page.goto('https://login.microsoftonline.com/b9203dbb-355b-4a6e-9ab3-b40f57429a4f/oauth2/authorize?client_id=00000003-0000-0ff1-ce00-000000000000&response_mode=form_post&response_type=code%20id_token&resource=00000003-0000-0ff1-ce00-000000000000&scope=openid&nonce=1A55C185A111A1B33E54CE141C818B87DA3FE0357864369A-FA4DFBA29CBE5E67FA15E4BE3235F5097820AC8A15FF643E426A67ECD8019884&redirect_uri=https%3A%2F%2Fkaizentechsa%2Esharepoint%2Ecom%2F_forms%2Fdefault%2Easpx&state=OD0w&claims=%7B"id_token"%3A%7B"xms_cc"%3A%7B"values"%3A%5B"CP1"%5D%7D%7D%7D&wsucxt=1&cobrandid=11bd8083-87e0-41b5-bb78-0bc43c8a8e8a&client-request-id=be9220a1-001f-8000-af7b-8a8c535dc2f7');
//   await page.fill('#i0116', 'lmonteiro@kaizen.tech');
//   await page.click('#idSIButton9');
//   await page.fill('#i0118', 'Lof25912'); 
//   await page.click('#idSIButton9');
//   await page.click('#idBtn_Back');

//   await page.context().storageState({ path: authFile });
// });

/*
setup('logicON', async ({ page }) => {

  await page.goto('http://ktmesapp02/TS/pages/acc/admin/services/logic$1/');
  await page.getByText('Start', {exact :true}).click();
  await page.getByRole('button', { name: 'OK'}).click();
  await page.waitForURL('http://ktmesapp02/TS/pages/root/admin/services/module$1/?c=ETS.Application.Wait.StatusActivityComplete&guid=5041653d-c63e-4706-a102-d31375233aa4');
});

setup('DMSon', async ({ page }) => {

  await page.goto('http://ktmesapp02/TS/pages/acc/admin/services/module$1/');
  await page.getByText('Start', {exact :true}).click();
  await page.getByRole('button', { name: 'OK'}).click();
  await page.waitForURL('http://ktmesapp02/TS/pages/acc/admin/services/module$1/?c=ETS.Application.Wait.StatusActivityComplete&guid=8d269f3a-4c61-432c-a12e-89a4be823889');
});*/