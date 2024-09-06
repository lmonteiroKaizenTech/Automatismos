import { test, expect } from '@playwright/test';
import { fail } from 'assert';
const TEMP = 60000;

test('Versões', async ({ page }) => {

    //test.setTimeout(TEMP)
    await page.setViewportSize({ width: 1920, height: 1000 });
    // #region Login
    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Verses/AllItems.aspx'); 
    await page.fill('#i0116', 'lmonteiro@kaizen.tech');
    await page.click('#idSIButton9');
    await page.fill('#i0118', 'Lof25912');
    await page.click('#idSIButton9');
    await page.click('#idBtn_Back');
    await page.waitForTimeout(2000);
    // #endregion

    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Tarefas%20de%20Distribuio%20de%20Verses/AllItems.aspx');
    await page.waitForTimeout(2000);
    await page.getByRole('menuitem').nth(1).click();
    await page.waitForTimeout(3000);
    await page.locator('cf-action-wrapper button').click();

    for (var i = 0; i < 20; i++)
        {
            await page.waitForTimeout(15000);
            await page.keyboard.press('a');
            await page.getByPlaceholder('Escreva para filtrar').type('teste');
            await page.keyboard.press('Enter');
            await page.locator('.field-C_x00f3_digo_x0020_Distribui_x00-htmlGrid_1').nth(-2).click({ force: true });
            await page.waitForTimeout(3000);
            await page.keyboard.press('a');
            await page.getByPlaceholder('Escreva para filtrar').type('teste');
            await page.keyboard.press('Enter');
            await page.waitForTimeout(3000);
            await page.locator('.field-Respons_x00e1_vel-htmlGrid_1').nth(-2).click({ force: true });
            await page.waitForTimeout(3000);
            await page.keyboard.press('a');
            await page.waitForTimeout(2000);
            await page.fill('.ms-BasePicker-input', 'Joana');
            await page.waitForTimeout(2000);
            await page.keyboard.press('Enter');
            await page.waitForTimeout(1000);
            await page.click(`#NewRowPlaceholderID`);
        }

    // await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Tarefas%20de%20Verso/AllItems.aspx');
    // await page.waitForTimeout(2000);

    // await page.getByRole('button').nth(2).click();
    // await page.waitForTimeout(10000);

    // await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Tarefas%20de%20Verso/AllItems.aspx');
    // await page.waitForTimeout(2000);

    // await page.click(`span:text("Código Versão")`);
    // await page.waitForTimeout(3000);
    // await page.click(`span:text("Filtrar por")`);
    // await page.click(`span:text("KTCore.3.4")`);

    // await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Tarefas%20de%20Distribuio%20de%20Verses/AllItems.aspx');
    // await page.waitForTimeout(2000);
    // await page.getByRole('menuitem').nth(1).click();
    // await page.waitForTimeout(3000);
    // await page.locator('cf-action-wrapper button').click();
    // await page.waitForTimeout(15000);
    // await page.keyboard.press('a');
    // await page.getByPlaceholder('Escreva para filtrar').type('#');
    // await page.keyboard.press('Enter');
    // await page.locator('.field-C_x00f3_digo_x0020_Distribui_x00-htmlGrid_1').nth(-2).click({ force: true });
    // await page.waitForTimeout(3000);
    // await page.keyboard.press('a');
    // await page.getByPlaceholder('Escreva para filtrar').type('AC');
    // await page.keyboard.press('Enter');
    // await page.waitForTimeout(3000);
    // await page.locator('.field-Respons_x00e1_vel-htmlGrid_1').nth(-2).click({ force: true });
    // await page.waitForTimeout(3000);
    // await page.keyboard.press('a');
    // await page.waitForTimeout(2000);
    // await page.fill('.ms-BasePicker-input','joana');
    // await page.waitForTimeout(2000);
    // await page.keyboard.press('Enter');

    // await page.getByRole('menuitem').nth(1).click();
    // await page.waitForTimeout(1000);
    // await page.click('#NewRowPlaceholderID');
    
    await page.close();

});