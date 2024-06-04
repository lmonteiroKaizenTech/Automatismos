import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';
const { exec } = require('child_process');

// Configurações de conexão
const sql = require('mssql');
const config = require('../../../AUTOMATISMOS/tests/dbConnection/connection.js');

let ambientes_nome: any[] = ['AC_PRD','AC_QLD','AC_TST','AFL_PRD','AFL_QLD','AFL_TST','ACF_PRD','ACF_QLD','ACF_TST','ACC_PRD','ACC_QLD','ACC_TST','DEV','AQS_PRD','AQS_TST','ARC_PRD','ARC_TST','ACO_PRD','ACO_TST','CLP_PRD','CLP_TST','DISNEYLAND', 'MCS_PRD','MCS_TST','FSL_PRD'];
let ambientes_links: any[] = ['AMR-MES15','AMRMMES89','ktmesapp04','AMR-MES16','AMRMMES88','KTMESAPP03','AMRMMES28','AMRMMES87','KTMESAPP05','AMRMMES30','AMRMMES84','ktmesapp02','ktmesapp01','62.28.206.115','172.16.1.15','172.16.3.1','172.16.1.9','172.16.4.2','172.16.1.13','10.60.101.20','ktmesapp07','ktdisneyland01','10.10.10.2','172.16.1.10','172.16.6.2'];

let output, user = '';
// Executar um comando PowerShell e capturar a saída
exec('chcp 65001 && whoami', (error, stdout, stderr) => {
    if (error) {
        console.error(`Erro ao executar o comando: ${error.message}`);
        return;
    }
    if (stderr) {
        console.error(`Erro do PowerShell: ${stderr}`);
        return;
    }

    // Faça o que quiser com a saída, como armazená-la em uma variável ou vetor
    output = stdout.trim(); // Remove espaços em branco extras

    if (output)
    {
        for (var i = 0; i < output.length; i++)
        {
            if (output[i] == '\\')
            {
                for (var j = 1; j < output.length - i; j++) user += output[i + j];
            }
        }
    }
});
test('GerarKey', async ({ page }) => {

    // -----------------Insira a Quantidade Aqui-----------------
    var quantidade = 3;
    // -----------------Insira a Quantidade Aqui-----------------

    let key, key_final;
    let keys_vetor: any[] = [];

    for (var i = 0; i < quantidade; i++)
    {
        // ------------------------------Gerar Key's------------------------------

        await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');
        const currentURL = page.url();
        if (currentURL.includes('http://ktmesapp01/TS/Account/LogOn.aspx?'))
        {
            await page.getByLabel('Login').fill('kt0032'); //utilizador kt
            await page.getByLabel('Password').click();
            await page.getByLabel('Password').fill('12345'); //password
            await page.getByRole('button', { name: 'Sign In' }).click();
        }

        await page.click('#contentPage_ctl32');
        await page.click('.btn-item-key-btn_GerarKey');
        await page.waitForTimeout(3000);
        key = await page.locator('#contentPage_ctl04').textContent();
        key_final = key?.trim();
        keys_vetor.push(key_final);

        // ------------------------------Fim Gerar Key's------------------------------
    }

    console.log('-----------------Keys Geradas-----------------');
    console.log(keys_vetor);
    console.log('-----------------Keys Geradas-----------------');

    await page.close();

});