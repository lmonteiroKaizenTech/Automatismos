import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';
const { exec } = require('child_process');

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
test('CriarTempos', async ({ page }) => {

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
    const dadosExcel = lerArquivoExcel('C:\\Users\\' + user + '\\Desktop\\CriarTempos.xlsx');
    console.log(dadosExcel);

    var linha = 0;
    let data: any[] = [];
    let colaborador: any[] = [];
    var item: any[] = [];
    var esforco: any[] = [];
    var incidente: any[] = [];

    for (var i = 0; i < dadosExcel.length; i++)
    {
        const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;
        data.push(segundaLinha['Data'] as string);
        colaborador.push(segundaLinha['Colaborador'] as string);
        item.push(segundaLinha['Item'] as string);
        esforco.push(segundaLinha['Esforço'] as string);
        incidente.push(segundaLinha['Incidente'] as string);
    }

    console.log(data);
    console.log(colaborador);
    console.log(item);
    console.log(esforco);
    console.log(incidente);

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

    await page.goto('https://kaizentechsa.sharepoint.com/sites/oper/s1/Lists/Tempo%202024/AllItems.aspx');
    await page.waitForTimeout(3000);
    for (var i = 0; i < dadosExcel.length; i++)
    {
        await page.getByTitle('Criar um novo item de lista nesta localização').click();
        await page.fill('#sp-DateTimePicker-DateTimeTextField', data[i]);
        await page.waitForTimeout(3000);
        await page.getByPlaceholder('Selecione uma opção').fill(colaborador[i]);
        await page.waitForTimeout(3000);
        await page.keyboard.press('Enter');
        await page.waitForTimeout(3000);
        await page.getByPlaceholder('Esta é uma coluna de pesquisa que apresenta dados de outra lista que excede atualmente o limiar de Vista de Lista definido pelo administrador (5000).').fill(item[i]);
        await page.waitForTimeout(3000);
        await page.keyboard.press('Enter');
        await page.waitForTimeout(3000);
        await page.getByPlaceholder('Introduza um número').fill(esforco[i]);
        await page.waitForTimeout(3000);
        if (incidente[i] != null)
        {
            const inc = await page.getByPlaceholder('Introduza o valor aqui').nth(2);
            const va = inc.first();
            va.fill(incidente[i]);
            await page.waitForTimeout(3000);
        }
        await page.click(`span:has-text("Guardar")`);
        await page.waitForTimeout(3000);
    }

    await page.close();

});