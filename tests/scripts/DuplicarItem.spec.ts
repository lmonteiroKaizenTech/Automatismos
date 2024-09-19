import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';
const { exec } = require('child_process');

test('DuplicarItem', async ({ page }) => {

    var item = 18324;

    await page.goto('https://login.microsoftonline.com/b9203dbb-355b-4a6e-9ab3-b40f57429a4f/oauth2/authorize?client_id=00000003-0000-0ff1-ce00-000000000000&response_mode=form_post&response_type=code%20id_token&resource=00000003-0000-0ff1-ce00-000000000000&scope=openid&nonce=1A55C185A111A1B33E54CE141C818B87DA3FE0357864369A-FA4DFBA29CBE5E67FA15E4BE3235F5097820AC8A15FF643E426A67ECD8019884&redirect_uri=https%3A%2F%2Fkaizentechsa%2Esharepoint%2Ecom%2F_forms%2Fdefault%2Easpx&state=OD0w&claims=%7B"id_token"%3A%7B"xms_cc"%3A%7B"values"%3A%5B"CP1"%5D%7D%7D%7D&wsucxt=1&cobrandid=11bd8083-87e0-41b5-bb78-0bc43c8a8e8a&client-request-id=be9220a1-001f-8000-af7b-8a8c535dc2f7');

    // const currentURL = page.url();
    // if (currentURL.includes('https://login.microsoftonline.com/b9203dbb-355b-4a6e-9ab3-b40f57429a4f/oauth2/authorize?client%5Fid=00000003%2D0000%2D0ff1%2Dce00%2D000000000000&response%5Fmode=form%5Fpost&response%5Ftype=code%20id%5Ftoken&resource=00000003%2D0000%2D0ff1%2Dce00%2D000000000000&scope=openid&nonce=EB8B2CD99ABC5D8851E81FA7126EF7BC366ADEA158DE1200%2DA0FCDA5195ECCD2A83F2B5F23A3A5DF7CD128C1F1436EBF6E1EC15D07465AB80&redirect%5Furi=https%3A%2F%2Fkaizentechsa%2Esharepoint%2Ecom%2F%5Fforms%2Fdefault%2Easpx&state=OD0w&claims=%7B%22id%5Ftoken%22%3A%7B%22xms%5Fcc%22%3A%7B%22values%22%3A%5B%22CP1%22%5D%7D%7D%7D&wsucxt=1&cobrandid=11bd8083%2D87e0%2D41b5%2Dbb78%2D0bc43c8a8e8a&client%2Drequest%2Did=d40c3fa1%2D30da%2D9000%2D6d44%2Daf9f0e1fa72a'))
    // {
        await page.fill('#i0116', 'lmonteiro@kaizen.tech');
        await page.click('#idSIButton9');
        await page.fill('#i0118', 'Lof25912');
        await page.click('#idSIButton9');
        await page.click('#idBtn_Back');
        await page.waitForTimeout(5000);
    // }
    // else if (currentURL.includes('https://login.microsoftonline.com/b9203dbb-355b-4a6e-9ab3-b40f57429a4f/oauth2/authorize?client_id=00000003-0000-0ff1-ce00-000000000000&response_mode=form_post&response_type=code%20id_token&resource=00000003-0000-0ff1-ce00-000000000000&scope=openid&nonce=1A55C185A111A1B33E54CE141C818B87DA3FE0357864369A-FA4DFBA29CBE5E67FA15E4BE3235F5097820AC8A15FF643E426A67ECD8019884&redirect_uri=https%3A%2F%2Fkaizentechsa%2Esharepoint%2Ecom%2F_forms%2Fdefault%2Easpx&state=OD0w&claims=%7B%22id_token%22%3A%7B%22xms'))
    // {
    //     const tile = await page.locator('.tile-container').first();
    //     tile.click();
    //     await page.fill('#i0118','Lof25912');
    //     await page.click('#idSIButton9');
    // }

    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Changes%20Follow%20Up/EditForm.aspx?ID=' + item + '&Source=%2Fsites%2Foper%2Fs1%2FLists%2FChanges%20Follow%20Up');
    
    await page.waitForTimeout(3000);

    let title = await page.locator('[id="Title_fa564e0f-0c70-4ab9-b863-0177e6ddd247_$TextField"]').inputValue();
    var relatable_item = await page.locator('[id="Item_x0020_relacionado_78ab82c8-67f3-483c-9dea-846f857b3e81_$TextField"]').inputValue();
    let project = await page.locator('[id="Project_13b20ece-43ac-4be9-b233-ab0804ecd635_$DropDownChoice"] option[selected="selected"]').textContent();
    let responsable = await page.locator('[id="Respons_x00e1_vel_7e214bdc-6cac-4eb8-a55b-30481ded226d_$ClientPeoplePicker_ResolvedList"]').textContent();
    let responsable_slice = responsable?.slice(0, -1);
    let begin_planed_date = await page.locator('[id="Plano_x002d_Data_x0020_In_x00ed__78d4f2eb-691c-45c9-a38d-207546dfcbb2_$DateTimeFieldDate"]').inputValue();
    let end_planed_date = await page.locator('[id="Plano_x002d_Data_x0020_Fim_860931a7-7052-412e-8f43-27d2bcbfdfc4_$DateTimeFieldDate"]').inputValue();
    let atribution = await page.locator('[id="AssignedTo_53101f38-dd2e-458c-b245-0c236cc13d1a_$ClientPeoplePicker_ResolvedList"]').textContent();
    let atribution_slice = atribution?.slice(0, -1);
    let state = await page.locator('[id="Status_3f277a5c-c7ae-4bbe-9d44-0456fb548f94_$DropDownChoice"] option[selected="selected"]').textContent();
    let wave = await page.locator('[id="Wave_x0020__x002d__x0020_Week_4e2d2350-7077-4f58-a51c-fda2df028639_$LookupField"] option[selected="selected"]').textContent();
    let sprint = await page.locator('[id="Sprint_5e95f44b-08d8-44cf-ab37-9c78e08954ce_$DropDownChoice"] option[selected="selected"]').textContent();
    var priority = await page.locator('[id="Priority_x0020_MES_6077134d-51f0-4c65-9471-a5bb194e175b_$NumberField"]').textContent();
    let Category = await page.locator('[id="Category_6df9bd52-550e-4a30-bc31-a4366832a87d_$DropDownChoice"] option[selected="selected"]').textContent();
    var effort_MES = await page.locator('[id="Est_x002e_Esfor_x00e7_o_x0020_ME_0426f208-b86a-4332-833d-4f3f69480dc6_$NumberField"]').inputValue();
    var execution = await page.locator('[id="Execu_x00e7__x00e3_o_x0020__x002_afa91fad-1d1d-47f1-99ab-7e52cf9436b8_$NumberField"]').textContent();
    let description = await page.locator('[id="Comment_6df9bd52-550e-4a30-bc31-a4366832a87f_$TextField_inplacerte"]').textContent();
    let Tested_Done_Date = await page.locator('[id="Tested_x0020_Done_x0020_Date_22abe953-52ee-43d3-b396-d316600f1a09_$DateTimeFieldDate"]').textContent();
    let analysis = await page.locator('[id="Analysis_2958527b-d669-4aba-99fd-b112ee969ce2_$TextField_inplacerte"]').textContent();
    let analyzed_by = await page.locator('[id="An_x00e1_lise_x0020_por_71a6c5df-218e-4629-8ee6-17c2ee784e54_$DropDownChoice"] option[selected="selected"]').textContent();
    let ticket = await page.locator('[id="Ticket_4a08fff8-c768-4e75-a197-285985f43de4_$TextField"]').textContent();
    let dev_done_date = await page.locator('[id="Dev_x0020_Done_x0020_Date_c0bc4b8a-b35b-4492-a871-66ffd9261c4b_$DateTimeFieldDate"]').textContent();
    let analysis_person = await page.locator('[id="Analysis_x0020_Person_01d8896a-218d-4303-982c-f2213057623e_$ClientPeoplePicker_ResolvedList"]').textContent();
    let analysis_person_slice = analysis_person?.slice(0, -1);
    console.log(begin_planed_date);
    console.log(end_planed_date);
    // console.log(title);
    // console.log(relatable_item);
    // console.log('Project: ' + project);
    // console.log(responsable);
    // console.log(atribution);
    // console.log(wave);
    // console.log(sprint);
    // console.log(priority);
    // console.log(Category);
    // console.log(esforco);
    // console.log(execution);
    // console.log(description);
    // console.log(analysis);
    // console.log(analyzed_by);
    // console.log(ticket);
    // console.log(dev_done_date);
    // console.log(analysis_person);

    await page.waitForTimeout(3000);

    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Changes%20Follow%20Up/Items%20Ativos1.aspx');
    await page.waitForTimeout(3000);

    await page.click('#idHomePageNewItem');

    if (title) await page.locator('[id="Title_fa564e0f-0c70-4ab9-b863-0177e6ddd247_$TextField"]').fill(title);
    if (relatable_item) await page.locator('[id="Item_x0020_relacionado_78ab82c8-67f3-483c-9dea-846f857b3e81_$TextField"]').fill(relatable_item);
    if (project) await page.locator('[id="Project_13b20ece-43ac-4be9-b233-ab0804ecd635_$DropDownChoice"]').selectOption(project);
    if (responsable_slice) await page.locator('[id="Respons_x00e1_vel_7e214bdc-6cac-4eb8-a55b-30481ded226d_$ClientPeoplePicker_EditorInput"]').fill(responsable_slice);
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    if (begin_planed_date) await page.locator('[id="Plano_x002d_Data_x0020_In_x00ed__78d4f2eb-691c-45c9-a38d-207546dfcbb2_$DateTimeFieldDate"]').fill(begin_planed_date);
    if (end_planed_date) await page.locator('[id="Plano_x002d_Data_x0020_Fim_860931a7-7052-412e-8f43-27d2bcbfdfc4_$DateTimeFieldDate"]').fill(end_planed_date);
    if (atribution_slice) await page.locator('[id="AssignedTo_53101f38-dd2e-458c-b245-0c236cc13d1a_$ClientPeoplePicker_EditorInput"]').fill(atribution_slice);
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    if (state) await page.locator('[id="Status_3f277a5c-c7ae-4bbe-9d44-0456fb548f94_$DropDownChoice"]').selectOption(state);
    if (wave) await page.locator('[id="Wave_x0020__x002d__x0020_Week_4e2d2350-7077-4f58-a51c-fda2df028639_$LookupField"]').selectOption(wave);
    if (sprint) await page.locator('[id="Sprint_5e95f44b-08d8-44cf-ab37-9c78e08954ce_$DropDownChoice"]').selectOption(sprint);
    if (priority) await page.locator('[id="Priority_x0020_MES_6077134d-51f0-4c65-9471-a5bb194e175b_$NumberField"]').fill(priority);
    if (Category) await page.locator('[id="Category_6df9bd52-550e-4a30-bc31-a4366832a87d_$DropDownChoice"]').selectOption(Category);
    if (effort_MES) await page.locator('[id="Est_x002e_Esfor_x00e7_o_x0020_ME_0426f208-b86a-4332-833d-4f3f69480dc6_$NumberField"]').fill(effort_MES);
    if (execution) await page.locator('[id="Execu_x00e7__x00e3_o_x0020__x002_afa91fad-1d1d-47f1-99ab-7e52cf9436b8_$NumberField"]').fill(execution);
    if (Tested_Done_Date) await page.locator('[id="Tested_x0020_Done_x0020_Date_22abe953-52ee-43d3-b396-d316600f1a09_$DateTimeFieldDate"]').fill(Tested_Done_Date);
    if (description) await page.locator('[id="Comment_6df9bd52-550e-4a30-bc31-a4366832a87f_$TextField_inplacerte"]').fill(description);
    if (analysis) await page.locator('[id="Analysis_2958527b-d669-4aba-99fd-b112ee969ce2_$TextField_inplacerte"]').fill(analysis);
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    if (analyzed_by) await page.locator('[id="An_x00e1_lise_x0020_por_71a6c5df-218e-4629-8ee6-17c2ee784e54_$DropDownChoice"]').selectOption(analyzed_by);
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    if (ticket) await page.locator('[id="Ticket_4a08fff8-c768-4e75-a197-285985f43de4_$TextField"]').fill(ticket);
    if (dev_done_date) await page.locator('[id="Dev_x0020_Done_x0020_Date_c0bc4b8a-b35b-4492-a871-66ffd9261c4b_$DateTimeFieldDate"]').fill(dev_done_date);
    if (analysis_person_slice) await page.locator('[id="Analysis_x0020_Person_01d8896a-218d-4303-982c-f2213057623e_$ClientPeoplePicker_EditorInput"]').fill(analysis_person_slice);
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.click('#ctl00_ctl30_g_0021adda_ba9e_4c0d_90a3_ccba980450b0_ctl00_toolBarTbl_RightRptControls_ctl00_ctl00_diidIOSaveItem');
    // await page.close();

});