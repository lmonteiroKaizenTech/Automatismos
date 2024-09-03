import { test, expect } from '@playwright/test';
import { fail } from 'assert';
const TEMP = 60000;

test('Versões', async ({ page }) => {

    //test.setTimeout(TEMP)

    // #region Variáveis
    let ambiente, responsavel: boolean = true;
    var aux;
    let titulos: any[] = [];
    // #endregion

    // #region Login
    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Verses/AllItems.aspx'); 
    await page.fill('#i0116', 'lmonteiro@kaizen.tech');
    await page.click('#idSIButton9');
    await page.fill('#i0118', 'Lof25912');
    await page.click('#idSIButton9');
    await page.click('#idBtn_Back');
    await page.waitForTimeout(2000);
    // #endregion

    // #region Versão
    await page.getByTitle('Título').click();
    await page.getByRole('menuitemcheckbox').nth(2).click();
    await page.waitForTimeout(2000);
    await page.getByPlaceholder('Escreva texto para localizar um filtro').fill('KTCore');
    await page.waitForTimeout(2000);
    const versao = await page.locator('.ms-Suggestions-container .ms-Button').first().textContent();
    // #endregion

    // #region Criar Distribuições de Versão
    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Distribuies%20de%20Verses/AllItems.aspx');
    await page.waitForTimeout(2000);

    for (var i = 0; i < 14; i++)
    {
        await page.getByRole('menuitem').first().click();
        switch (i) {
            case 0:
                ambiente = 'AC_PRD_KTCore';
                break;
            case 1:
                ambiente = 'AC_QLD_KTCore';
                break;
            case 2:
                ambiente = 'ACC_PRD_KTCore';
                responsavel = false;
                break;
            case 3:
                ambiente = 'ACC_QLD_KTCore';
                responsavel = false;
                break;
            case 4:
                ambiente = 'ACF_PRD_KTCore';
                break;
            case 5:
                ambiente = 'ACF_QLD_KTCore';
                break;
            case 6:
                ambiente = 'ACO_PRD_KTCore';
                break;
            case 7:
                ambiente = 'AFL_PRD_KTCore';
                responsavel = false;
                break;
            case 8:
                ambiente = 'AFL_QLD_KTCore';
                responsavel = false;
                break;
            case 9:
                ambiente = 'AQS_PRD_KTCore';
                responsavel = false;
                break;
            case 10:
                ambiente = 'ARC_PRD_KTCore';
                break;
            case 11:
                ambiente = 'CLP_PRD_KTCore';
                break;
            case 12:
                ambiente = 'FSL_PRD_KTCore';
                break;
            case 13:
                ambiente = 'PCL_PRD_KTCore';
                break;
            default:
                break;
        }
        if (versao)
        {
            await page.getByPlaceholder('Introduza o valor aqui').first().fill(ambiente + versao);
            await page.getByPlaceholder('Selecione uma opção').first().fill(ambiente);
            await page.waitForTimeout(2000);
            const Suggestions1 = await page.locator('.ms-Suggestions-container .ms-Button').first().textContent();
            if (!Suggestions1?.includes(ambiente)) await page.getByPlaceholder('Selecione uma opção').first().fill(ambiente); await page.waitForTimeout(1000);
            await page.keyboard.press('Enter');
            await page.getByPlaceholder('Selecione uma opção').fill(versao);
            await page.waitForTimeout(2000);
            const Suggestions2 = await page.locator('.ms-Suggestions-container .ms-Button').first().textContent();
            if (!Suggestions2?.includes(versao)) await page.getByPlaceholder('Selecione uma opção').first().fill(versao); await page.waitForTimeout(1000);
            await page.keyboard.press('Enter');
            await page.locator('.ReactClientForm-editButtons .ms-Button').first().click();
            await page.waitForTimeout(2000);
        }
    }
    // #endregion

    // #region Tarefas de Versão

    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Tarefas%20de%20Verso/AllItems.aspx');
    await page.waitForTimeout(2000);

    const titulo = await page.$$('[role="gridcell"] [role="button"]');

    for (var i = 0; i < titulo.length; i++)
    {
        aux = await titulo[i].textContent();
        titulos.push(aux);
    }

    // #endregion

    // #region Tarefas de Distribuição

    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Tarefas%20de%20Distribuio%20de%20Verses/AllItems.aspx');
    await page.waitForTimeout(2000);

    await page.getByRole('menuitem').first().click();
    await page.getByPlaceholder('Selecione uma opção').first().fill(ambiente + versao);
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.getByPlaceholder('Selecione uma opção').nth(1).type(titulos[0]);
    // await page.waitForSelector(selector);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);
    if (responsavel) await page.getByPlaceholder('Introduza um nome ou endereço de e-mail').fill('Ines tomas');
    else await page.getByPlaceholder('Introduza um nome ou endereço de e-mail').fill('joana carvalho');

    // #endregion

    await page.close();

});