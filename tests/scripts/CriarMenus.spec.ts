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
test('CriarMenus', async ({ page }) => {

    // Definindo o tipo de uma linha do Excel
    type LinhaExcel = Record<string, string | null>;

    // Função para ler o arquivo Excel
    function lerArquivoExcel(nomeArquivo: string): LinhaExcel[] {
        // Carrega o arquivo
        const workbook = XLSX.readFile(nomeArquivo);

        // Pega a primeira planilha do arquivo
        const primeiraPlanilha = workbook.Sheets[workbook.SheetNames[0]];

        // Converte os dados da planilha em um objeto JSON
        const dados = XLSX.utils.sheet_to_json(primeiraPlanilha, { header: 1 }) as string[][];

        // Extrai os cabeçalhos da primeira linha
        const colunas = dados[0];

        // Inicializa um array para armazenar os dados
        const dadosFormatados: LinhaExcel[] = [];

        // Itera sobre as linhas de dados, começando da segunda linha
        for (let i = 1; i < dados.length; i++) {
            const linha: LinhaExcel = {};
            // Itera sobre as colunas
            for (let j = 0; j < colunas.length; j++) {
                const valor = dados[i][j];
                linha[colunas[j]] = valor !== undefined ? valor.toString() : null;
            }
            dadosFormatados.push(linha);
        }

        // Retorna os dados formatados
        return dadosFormatados;
    }
    
    // Exemplo de uso
    let dadosExcel

    // Tentar ler o arquivo Excel com o caminho completo
    try {
        dadosExcel = lerArquivoExcel('C:\\Users\\' + user + '\\OneDrive - KAIZEN TECH, S.A\\Ambiente de Trabalho\\Páginas.xlsx');
        console.log("Leitura bem-sucedida com OneDrive.");
    } catch (error) {
        // Se ocorrer um erro, tentar ler o arquivo Excel com o caminho usando a variável 'user'
        try {
            dadosExcel = lerArquivoExcel('C:\\Users\\' + user + '\\Desktop\\Páginas.xlsx');
            console.log("Leitura bem-sucedida sem OneDrive.");
        } catch (error) {
            console.error("Erro ao ler o arquivo Excel:", error.message);
        }
    }

    console.log(dadosExcel);

    let ambiente;
    let site;

    const segundaLinha: LinhaExcel = dadosExcel[0] as LinhaExcel;

    ambiente = segundaLinha['Ambiente'] as string;
    site = segundaLinha['Site'] as string;

    let cpag: any[] = [];
    let key_vetor: any[] = [];
    let NovoNome: any[] = [];
    let visibilidade: any[] = [];

    for (var i = 0; i < dadosExcel.length; i++)
    {
        const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;
        cpag.push(segundaLinha['Cpag'] as string);
        key_vetor.push(segundaLinha['Key'] as string);
        NovoNome.push(segundaLinha['Novo Nome'] as string);
        visibilidade.push(segundaLinha['Visibilidade'] as string);
    }
    console.log('Key: ' + key_vetor);
    console.log('Novo nome: ' + NovoNome);
    console.log('Visibilidade: ' + visibilidade);

    console.log(cpag);

    let vetor: string[][] = [];
    var linha = 0;

    for (var k = 0; k < dadosExcel.length; k++)
    {
        let ma: any[] = [];
        var i = 1;
        let proximo = 'começo';
        while (proximo != null)
        {
            const segundaLinha: LinhaExcel = dadosExcel[linha] as LinhaExcel;
            
            const este = segundaLinha[i] as string;
            proximo = segundaLinha[i + 1] as string;
            ma.push(este);
            i++;
        }
        linha++;
        vetor.push(ma);
    }
    
    // i++;
    // const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;
    // const este = segundaLinha[i] as string;
    // vetor.push(este);
    console.log('Vetor: ' + vetor);

    var position = 0;
    for (var i = 0; i < ambientes_nome.length; i++)
    {
        if (ambiente == ambientes_nome[i]) position = i;
    }

    let ambiente_final;
    for (var i = 0; i < ambientes_links.length; i++)
    {
        if (i == position) ambiente_final = ambientes_links[i];
    }

    // ---------------Login Site Principal---------------
    // TrakSys12

    if (ambiente == 'AQS_PRD')
    {
        await page.goto('https://' + ambiente_final + '/TS/');
        await page.waitForTimeout(3000);
        //Verificação de Login
        const currentURL = page.url();
        if (currentURL == 'https://' + ambiente_final +'/TS/Account/LogOn.aspx?ts_deny=true&ts_rurl=%2fTS%2fdefault.aspx')
        {
            await page.getByLabel('Login').fill('kt0032'); //utilizador kt
            await page.getByLabel('Password').click();
            await page.getByLabel('Password').fill('Lof25912'); // password
            await page.getByRole('button', { name: 'Sign In' }).click();

        }
    }
    else
    {
        await page.goto('http://' + ambiente_final + '/TS/');
        await page.waitForTimeout(3000);
        //Verificação de Login
        const currentURL = page.url();
        if (currentURL == 'http://' + ambiente_final +'/TS/Account/LogOn.aspx?ts_deny=true&ts_rurl=%2fTS%2fdefault.aspx')
        {
            await page.getByLabel('Login').fill('kt0032'); //utilizador kt
            await page.getByLabel('Password').click();
            await page.getByLabel('Password').fill('12345'); // password
            await page.getByRole('button', { name: 'Sign In' }).click();
    
        }
    }

    var linha = 0;
    
    for (var j = 0; j < dadosExcel.length; j++)
    {
        let key, key_final;

        if (key_vetor[j] == null)
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

            // ------------------------------Fim Gerar Key's------------------------------
        }

        //let vetor_primeiro: any[] = [];
        var soma = 0;

        await page.waitForTimeout(5000);
        if (ambiente == 'AQS_PRD') await page.goto('https://' + ambiente_final + '/TS/pages/' + site + '/dev/pagedef/');
        else await page.goto('http://' + ambiente_final + '/TS/pages/' + site + '/dev/pagedef/');
        await page.waitForTimeout(3000);
        for (var i = 0; i < vetor[j].length; i++)
        {
            console.log('vetor: ' + vetor[j][i]);
            if (vetor[j][i].includes('Hub'))  await page.click(`a:text("${vetor[j][i]}")`);
            if (vetor[j][i].includes('Spokes'))
            {
                const segundo = await page.locator(`a:text("${vetor[j][i]}")`).nth(1);
                const eventhandler = segundo.first();
                eventhandler.click();
            }
            if (vetor[j][i] == 'New') await page.click('a:text("New")');
            else
            {
                const elementos = await page.$$(`[title="${vetor[j][i]}"]`);
                //vetor_primeiro.push(await page.getByTitle(vetor[j][i]).textContent());
                for (var k = 0; k < elementos.length; k++)
                {
                    const texto = await elementos[k].textContent();
                    if (texto == 'Hubs' || texto == 'Spokes') soma++;
                    await page.waitForTimeout(3000);
                    const texto_final = texto?.trim();
                    console.log('-------------------------------------------------------------');
                    console.log(texto_final);
                    console.log('-------------------------------------------------------------');
                    console.log('Texto: ' + texto_final);
                    if (texto_final == vetor[j][i]) elementos[k].click();
                }
            }
            await page.waitForTimeout(3000);
        }

        console.log('soma: ' + soma);
        // if (soma > 0)
        // {
        //     const pri = await page.locator(`a:has-text("Hub")`).nth(2);
        //     const EventHandler = pri.first();
        //     EventHandler.click();
        //     await page.waitForTimeout(3000);
        // }
        // else
        // {
        //     const pri = await page.locator(`a:has-text()`).first();
        //     pri.click();
        //     await page.waitForTimeout(3000);
        // }
        
        await page.waitForTimeout(3000);
        await page.click(`a:text("Import")`);

        await page.waitForTimeout(3000);

        // Localize o input de arquivo e insira o caminho do arquivo Excel
        const inputFile = await page.$('input[type="file"]');
        if (inputFile)
        {
            // Tentar ler o arquivo Excel com o caminho completo
            try {
                await inputFile.setInputFiles('C:\\Users\\' + user + '\\OneDrive - KAIZEN TECH, S.A\\Ambiente de Trabalho\\Paginas\\' + cpag[j]);
            } catch (error) {
                // Se ocorrer um erro, tentar ler o arquivo Excel com o caminho sem OneDrive
                try {
                    await inputFile.setInputFiles('C:\\Users\\' + user + '\\Desktop\\Paginas\\' + cpag[j]);

                } catch (error) {
                    console.error("Erro ao ler o arquivo Excel:", error.message);
                }
            }
        }
        await page.waitForTimeout(3000);

        if (key_vetor[j] != null) await page.fill('#contentPage_slice2_KeyTextBox', key_vetor[j]);
        else if (key) await page.fill('#contentPage_slice2_KeyTextBox', key);

        await page.waitForTimeout(3000);
        await page.click('#contentPage_slice2_Create');

        await page.waitForTimeout(5000);

        // ------------------------Parametrizar Página------------------------

        if (NovoNome[j] && visibilidade[j] == null)
        {
            if (key_vetor[j] == null)
            {
                const elementos = await page.$$(`[title*="${key_final}"]`);
                elementos[0].click();
            }
            else
            {
                const elementos = await page.$$(`[title*="${key_vetor[j]}"]`);
                elementos[0].click();
            }
            await page.waitForTimeout(3000);
            await page.locator('.fa-edit').nth(1).click();
            await page.waitForTimeout(3000);
            await page.fill('#tseditName', NovoNome[j]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
            await page.waitForTimeout(3000);
        }
        else if (visibilidade[j] && NovoNome[j])
        {
            if (key_vetor[j] == null)
            {
                const elementos = await page.$$(`[title*="${key_final}"]`);
                elementos[0].click();
            }
            else
            {
                const elementos = await page.$$(`[title*="${key_vetor[j]}"]`);
                elementos[0].click();
            }
            await page.waitForTimeout(3000);
            await page.locator('.fa-edit').nth(1).click();
            await page.waitForTimeout(3000);
            await page.fill('#tseditName', NovoNome[j]);
            await page.waitForTimeout(3000);
            await page.click('a:text("Visibility")');
            await page.waitForTimeout(3000);
            await page.click('#tseditIsVisibleInNavigation');
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
        }
        else if (visibilidade[j] && NovoNome[j] == null)
        {
            if (key_vetor[j] == null)
            {
                const elementos = await page.$$(`[title*="${key_final}"]`);
                elementos[0].click();
            }
            else
            {
                const elementos = await page.$$(`[title*="${key_vetor[j]}"]`);
                elementos[0].click();
            }
            
            await page.waitForTimeout(3000);
            await page.locator('.fa-edit').nth(1).click();
            await page.waitForTimeout(3000);
            await page.click('a:text("Visibility")');
            await page.waitForTimeout(3000);
            await page.click('#tseditIsVisibleInNavigation');
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
            await page.waitForTimeout(3000);
        }

        // ----------------------Fim Parametrizar Página----------------------

        linha++;
        console.log('-------------------------Linha do excel executada-------------------------');
        console.log(linha);
        console.log('--------------------------------------------------------------------------');
    }

    await page.close();

});