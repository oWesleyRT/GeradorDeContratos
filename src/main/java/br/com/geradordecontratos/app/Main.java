package br.com.geradordecontratos.app;


import br.com.geradordecontratos.GeradorDeContratos;

import javax.swing.*;

public class Main {
    public static void main(String[] args) {

        String nome = JOptionPane.showInputDialog("Digite o nome completo: ");
        String qualificacaoCivil = JOptionPane.showInputDialog("Digite a qualificação civil completa: ");
        String parcelas = JOptionPane.showInputDialog("Digite o numero de parcelas: ");


        GeradorDeContratos geradorDeContratos = new GeradorDeContratos();
        String data = geradorDeContratos.getData();
        //APOSENTADORIA
        geradorDeContratos.editarDocumentoAposentadoria(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfAposentadoria();
        //PLANEJAMENTO RGPS
        geradorDeContratos.editarDocumentoPlanejamento(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfPlanejamento();
        //CORREÇÃO CNIS INTEGRAL
        geradorDeContratos.editarDocumentoCorrecaoIntegral(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfCorrecaoIntegral();
        //CORREÇÃO CNIS DESCONTO
        geradorDeContratos.editarDocumentoCorrecaoDesconto(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfCorrecaoDesconto();
        //AVERBAÇÃO
        geradorDeContratos.editarDocumentoAverbacao(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfAverbacao();
        //RETENÇÃO 25%
        geradorDeContratos.editarDocumentoRetencao25(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfRetencao25();
        //LIMINAR 25%
        geradorDeContratos.editarDocumentoLiminar25(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfLiminar25();
        //REVISÃO CTC COM DESCONTO
        geradorDeContratos.editarDocumentoRevisaoCtcDesconto(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfRevisaoCtcDesconto();
        //REVISÃO RGPS
        geradorDeContratos.editarDocumentoRevisaoRgps(nome, qualificacaoCivil, data, parcelas);
        geradorDeContratos.converterDocxPdfRevisaoRgps();

        System.exit(0);

    }
}