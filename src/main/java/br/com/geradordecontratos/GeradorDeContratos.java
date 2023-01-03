package br.com.geradordecontratos;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;

public class GeradorDeContratos {

    public String getData(){
        Calendar calendar = Calendar.getInstance();
        Date dataAtualSemFormato = calendar.getTime();
        SimpleDateFormat sdf = new SimpleDateFormat("dd 'de' MMMM 'de' yyyy", new Locale("pt", "BR"));
        String dataAtual = sdf.format(dataAtualSemFormato);
        return dataAtual;
    }

    public void editarDocumentoAposentadoria(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/AposentadoriaRgps.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/AposentadoriaRgps.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfAposentadoria() {
        System.out.println("Iniciando a conversão contrato aposentadoria");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/AposentadoriaRgps.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/AposentadoriaRgps.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato aposentadoria gerado");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void editarDocumentoPlanejamento(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/PlanejamentoRgps.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/PlanejamentoRgps.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfPlanejamento() {
        System.out.println("Iniciando a conversão contrato planejamento");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/PlanejamentoRgps.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/PlanejamentoRgps.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato planejamento gerado");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void editarDocumentoCorrecaoIntegral(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/CorrecaoCnisIntegral.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/CorrecaoCnisIntegral.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfCorrecaoIntegral() {
        System.out.println("Iniciando a conversão contrato correção integral");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/CorrecaoCnisIntegral.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/CorrecaoCnisIntegral.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato correção integral gerado");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void editarDocumentoCorrecaoDesconto(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/CorrecaoCnisDesconto.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/CorrecaoCnisDesconto.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfCorrecaoDesconto() {
        System.out.println("Iniciando a conversão contrato correção desconto");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/CorrecaoCnisDesconto.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/CorrecaoCnisDesconto.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato correção desconto gerado");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void editarDocumentoAverbacao(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/Averbacao.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/Averbacao.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfAverbacao() {
        System.out.println("Iniciando a conversão contrato averbacao");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/Averbacao.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/Averbacao.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato averbacao gerado");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void editarDocumentoRetencao25(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/Retencao25.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/Retencao25.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfRetencao25() {
        System.out.println("Iniciando a conversão contrato retenção");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/Retencao25.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/Retencao25.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato retenção gerado");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void editarDocumentoRevisaoCtcDesconto(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/RevisaoCtcDesconto.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/RevisaoCtcDesconto.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfRevisaoCtcDesconto() {
        System.out.println("Iniciando a conversão contrato revisão CTC desconto");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/RevisaoCtcDesconto.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/RevisaoCtcDesconto.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato revisão CTC desconto gerado");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void editarDocumentoRevisaoRgps(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/RevisaoRgps.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/RevisaoRgps.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfRevisaoRgps() {
        System.out.println("Iniciando a conversão contrato revisão rgps");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/RevisaoRgps.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/RevisaoRgps.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato revisão rgps gerado");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void editarDocumentoLiminar25(String nome, String qualificacaoCivil, String dataAtual, String parcelas) {
        File outputFileEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/Liminar25.docx");
        try {
            // Abre o primeiro contrato
            File inputWordEditar = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/BASE/Liminar25.docx");
            FileInputStream inputStream = new FileInputStream(inputWordEditar);
            XWPFDocument documento = new XWPFDocument(inputStream);

            for (XWPFParagraph paragraph : documento.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null && text.contains("[NOME]")) {
                        text = text.replace("[NOME]", nome);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[QUALIFICACAOCIVIL]")) {
                        text = text.replace("[QUALIFICACAOCIVIL]", qualificacaoCivil);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[DATA]")) {
                        text = text.replace("[DATA]", dataAtual);
                        run.setText(text, 0);
                    }
                    if (text != null && text.contains("[PARCELAS]")) {
                        text = text.replace("[PARCELAS]", parcelas);
                        run.setText(text, 0);
                    }
                }
            }

            // Fecha o DOCX
            FileOutputStream outputStream = new FileOutputStream(outputFileEditar);
            documento.write(outputStream);
            outputStream.close();
            documento.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void converterDocxPdfLiminar25() {
        System.out.println("Iniciando a conversão contrato liminar 25");
        File inputWord = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/EDITADOS/Liminar25.docx");
        File outputFile = new File("C:/Users/PICHAU/Documents/01 - Koetz/CONTRATO HONORARIOS/GERADOR/PDF/Liminar25.pdf");
        try  {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
            System.out.println("Contrato revisão liminar 25");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
