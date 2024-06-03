import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';
const { exec } = require('child_process');

test('CriarItemSuporte', async ({ page }) => {

    await page.goto('https://login.microsoftonline.com/b9203dbb-355b-4a6e-9ab3-b40f57429a4f/oauth2/authorize?client_id=00000003-0000-0ff1-ce00-000000000000&response_mode=form_post&response_type=code%20id_token&resource=00000003-0000-0ff1-ce00-000000000000&scope=openid&nonce=1A55C185A111A1B33E54CE141C818B87DA3FE0357864369A-FA4DFBA29CBE5E67FA15E4BE3235F5097820AC8A15FF643E426A67ECD8019884&redirect_uri=https%3A%2F%2Fkaizentechsa%2Esharepoint%2Ecom%2F_forms%2Fdefault%2Easpx&state=OD0w&claims=%7B"id_token"%3A%7B"xms_cc"%3A%7B"values"%3A%5B"CP1"%5D%7D%7D%7D&wsucxt=1&cobrandid=11bd8083-87e0-41b5-bb78-0bc43c8a8e8a&client-request-id=be9220a1-001f-8000-af7b-8a8c535dc2f7');
    await page.waitForTimeout(3000);
    const currentURL = page.url();
    if (currentURL.includes('https://login.microsoftonline.com/b9203dbb-355b-4a6e-9ab3-b40f57429a4f/oauth2/authorize?client_id=00000003-0000-0ff1-ce00-000000000000&response_mode=form_post&response_type=code%20id_token&resource=00000003-0000-0ff1-ce00-000000000000&scope=openid&nonce=1A55C185A111A1B33E54CE141C818B87DA3FE0357864369A-FA4DFBA29CBE5E67FA15E4BE3235F5097820AC8A15FF643E426A67ECD8019884&redirect_uri=https%3A%2F%2Fkaizentechsa%2Esharepoint%2Ecom%2F_forms%2Fdefault%2Easpx&state=OD0w&claims=%7B"id_token"%3A%7B"xms_cc"'))
    {
        await page.fill('#i0116', 'lmonteiro@kaizen.tech');
        await page.click('#idSIButton9');
        await page.fill('#i0118', 'Lof25912'); 
        await page.click('#idSIButton9');
        await page.click('#idBtn_Back');
        await page.waitForTimeout(5000);
    }
    else if (currentURL.includes('https://login.microsoftonline.com/b9203dbb-355b-4a6e-9ab3-b40f57429a4f/oauth2/authorize?client_id=00000003-0000-0ff1-ce00-000000000000&response_mode=form_post&response_type=code%20id_token&resource=00000003-0000-0ff1-ce00-000000000000&scope=openid&nonce=1A55C185A111A1B33E54CE141C818B87DA3FE0357864369A-FA4DFBA29CBE5E67FA15E4BE3235F5097820AC8A15FF643E426A67ECD8019884&redirect_uri=https%3A%2F%2Fkaizentechsa%2Esharepoint%2Ecom%2F_forms%2Fdefault%2Easpx&state=OD0w&claims=%7B%22id_token%22%3A%7B%22xms'))
    {
        const tile = await page.locator('.tile-container').first();
        tile.click();
        await page.fill('#i0118','Lof25912');
        await page.click('#idSIButton9');
    }

    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Changes%20Follow%20Up/Items%20Ativos1.aspx');
    await page.waitForTimeout(3000);

    await page.click('#idHomePageNewItem');
    
    function getWeekNumber(date: Date): number {
        // Copying date so the original date won't be modified
        const tempDate = new Date(date.valueOf());
    
        // ISO week date weeks start on Monday, so correct the day number
        const dayNum = (date.getDay() + 6) % 7;
    
        // Set the target to the nearest Thursday (current date + 4 - current day number)
        tempDate.setDate(tempDate.getDate() - dayNum + 3);
    
        // ISO 8601 week number of the year for this date
        const firstThursday = tempDate.valueOf();
    
        // Set the target to the first day of the year
        // First set the target to January 1st
        tempDate.setMonth(0, 1);
    
        // If this is not a Thursday, set the target to the next Thursday
        if (tempDate.getDay() !== 4) {
            tempDate.setMonth(0, 1 + ((4 - tempDate.getDay()) + 7) % 7);
        }
    
        // The weeknumber is the number of weeks between the first Thursday of the year
        // and the Thursday in the target week
        return 1 + Math.ceil((firstThursday - tempDate.valueOf()) / 604800000); // 604800000 = number of milliseconds in a week
    }
    const currentDate = new Date();
    const semana = getWeekNumber(currentDate); 

    await page.fill('[id="Title_fa564e0f-0c70-4ab9-b863-0177e6ddd247_$TextField"]', 'Registo Tempo Suporte 04_2024 - Semana ' + semana);
    await page.waitForTimeout(3000);
    await page.selectOption('[id="Dossier_18e3039a-bbaa-4edf-a552-fcaf57d64898_$LookupField"]', '(60) Suporte');
    await page.waitForTimeout(3000);
    await page.selectOption('[id="Project_13b20ece-43ac-4be9-b233-ab0804ecd635_$DropDownChoice"]', 'SP.24.001 - Suporte Amorim');
    await page.waitForTimeout(3000);
    await page.fill('[id="Respons_x00e1_vel_7e214bdc-6cac-4eb8-a55b-30481ded226d_$ClientPeoplePicker_EditorInput"]', 'tiago costa');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.fill('[id="Analysis_x0020_Person_01d8896a-218d-4303-982c-f2213057623e_$ClientPeoplePicker_EditorInput"]', 'tiago costa');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.fill('[id="AssignedTo_53101f38-dd2e-458c-b245-0c236cc13d1a_$ClientPeoplePicker_EditorInput"]', 'tiago costa');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.fill('[id="Validation_x0020_Person_db3b4462-feb6-4d75-8302-0f30c9c37cfe_$ClientPeoplePicker_EditorInput"]', 'tiago costa');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.selectOption('[id="Status_3f277a5c-c7ae-4bbe-9d44-0456fb548f94_$DropDownChoice"]', '(2) In Development');
    await page.waitForTimeout(3000);
    await page.selectOption('[id="Wave_x0020__x002d__x0020_Week_4e2d2350-7077-4f58-a51c-fda2df028639_$LookupField"]', 'W24' + semana);
    await page.waitForTimeout(3000);
    await page.fill('[id="Priority_x0020_MES_6077134d-51f0-4c65-9471-a5bb194e175b_$NumberField"]', '0');
    await page.waitForTimeout(3000);
    await page.selectOption('[id="Category_6df9bd52-550e-4a30-bc31-a4366832a87d_$DropDownChoice"]', '(4) Task');
    await page.waitForTimeout(3000);
    await page.fill('[id="Comment_6df9bd52-550e-4a30-bc31-a4366832a87f_$TextField_inplacerte"]', 'Registo Tempo Suporte 04_2024 - Semana ' + semana);
    await page.waitForTimeout(3000);
    await page.click('#ctl00_ctl30_g_0021adda_ba9e_4c0d_90a3_ccba980450b0_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem');
    await page.waitForTimeout(3000);

    await page.close();

});