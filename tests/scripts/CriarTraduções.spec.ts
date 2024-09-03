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
test('CriarTraduções', async ({ page }) => {

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
        dadosExcel = lerArquivoExcel('C:\\Users\\' + user + '\\OneDrive - KAIZEN TECH, S.A\\Ambiente de Trabalho\\AtualizarTraducoes.xlsx');
        console.log("Leitura bem-sucedida com OneDrive.");
    } catch (error) {
        // Se ocorrer um erro, tentar ler o arquivo Excel com o caminho usando a variável 'user'
        try {
            dadosExcel = lerArquivoExcel('C:\\Users\\' + user + '\\Desktop\\AtualizarTraducoes.xlsx');
            console.log("Leitura bem-sucedida sem OneDrive.");
        } catch (error) {
            console.error("Erro ao ler o arquivo Excel:", error.message);
        }
    }

    console.log(dadosExcel);

    let key: any[] = [];
    let resource: any[] = [];

    for (var i = 0; i < dadosExcel.length; i++)
    {
        const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;
        key.push(segundaLinha['key'] as string);
        resource.push(segundaLinha['Resource'] as string);
    }
    console.log(key);

    console.log(resource);

    let vetor: string[][] = [];
    var linha = 0;
    
    console.log('Vetor: ' + vetor);


    // ---------------Login Site Principal---------------
    // TrakSys12

    await page.goto('http://172.16.1.17/TS');
    await page.getByLabel('Login').fill('kt0032'); //utilizador kt
    await page.getByLabel('Password').click();
    await page.getByLabel('Password').fill('1234'); // password
    await page.getByRole('button', { name: 'Sign In' }).click();

    var linha = 0;

    for (var i = 0; i < dadosExcel.length; i++)
    {
        await page.goto('http://172.16.1.17/TS/pages/fsl/config/resx/?S1ID=168&S2Key=List.ResourceGroup.Items&S2ParentID=168&S3Key=Item.ResourceGroup&S3ID=168&s1=11020');
        await page.waitForTimeout(3000);
        await page.locator('a:text("  New")').nth(1).click();
        await page.fill('#tseditKey', key[i]);
        await page.fill('#tseditDefaultValue', resource[i]);
        await page.waitForTimeout(3000);
        await page.click('#contentPage_SaveAndNew_Button');
    }

    linha++;
    console.log('-------------------------Linha do excel executada-------------------------');
    console.log(linha);
    console.log('--------------------------------------------------------------------------');


    await page.close();

});