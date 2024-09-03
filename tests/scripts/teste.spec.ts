import { test, expect } from '@playwright/test';
import { fail } from 'assert';
const TEMP = 60000;

test('Versões', async ({ page }) => {

    //test.setTimeout(TEMP)
    // #region Login
    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Verses/AllItems.aspx'); 
    await page.fill('#i0116', 'lmonteiro@kaizen.tech');
    await page.click('#idSIButton9');
    await page.fill('#i0118', 'Lof25912');
    await page.click('#idSIButton9');
    await page.click('#idBtn_Back');
    await page.waitForTimeout(2000);
    // #endregion

    function teste() {
        
    }

    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Tarefas%20de%20Distribuio%20de%20Verses/AllItems.aspx');
    await page.waitForTimeout(2000);

    await page.getByRole('menuitem').first().click();
    await page.getByPlaceholder('Selecione uma opção').first().fill('Teste');
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.getByPlaceholder('Selecione uma opção').nth(1).type('', { delay: 100 });

    await page.waitForTimeout(10000);
    // await page.waitForSelector(selector);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    await page.getByPlaceholder('Introduza um nome ou endereço de e-mail').fill('Ines tomas');
    await page.waitForSelector('');

    await page.waitForTimeout(10000);

    await page.close();

});