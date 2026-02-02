package br.com.cpqd;

import static org.junit.Assert.assertTrue; 

import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.util.Calendar;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


public class WebServicesTest {
	private String workdir;
	private String user;
	private String password;
	private Workbook inputRT;
	private WritableWorkbook outputRT;
	private Sheet inputSheet;
	private Cell cell[];
	private Calendar calendar;
	private String date;
	private String url;
	private String userGP [];	
	
	@Before																												//Sera executado antes do teste (chamada do metodo testSequence)
	public void setup() throws BiffException, IOException {
		WorkbookSettings ws = new WorkbookSettings();
		ws.setEncoding("Cp1252");													
		calendar = Calendar.getInstance();
		date = Integer.toString(calendar.get(Calendar.DAY_OF_MONTH)) + "-"
				+ Integer.toString(calendar.get(Calendar.MONTH) + 1) + "-"
				+ Integer.toString(calendar.get(Calendar.YEAR)) + "_"
				+ Integer.toString(calendar.get(Calendar.HOUR_OF_DAY)) + "-"
				+ Integer.toString(calendar.get(Calendar.MINUTE)) + "-"
				+ Integer.toString(calendar.get(Calendar.SECOND));

		inputRT = Workbook.getWorkbook(new File("RT-WebServices_Vivo_Teste_Automatico.xls"),ws);
		inputSheet = inputRT.getSheet(0); 																				// A planilha 0 possui dados de configuracao
		cell = inputSheet.getRow(1);
		url = cell[0].getContents();
		user = cell[1].getContents();
		password = cell[2].getContents();
		workdir = "target";																								// Diretorio target criado pelo Maven e onde serao armazenados o arquivo de log e tambem o roteiro com os dados da execucao
		userGP = user.split("/");	
		outputRT = Workbook.createWorkbook(new File(workdir + "/RT-WebServices_Vivo_Teste_Automatico_executado_" + userGP[1] + "_" + date + ".xls"),	inputRT, ws);
	}

	@After																												//Sera executado apos o teste (chamada do metodo testSequence)
	public void tearDown() throws IOException, WriteException {
		inputRT.close();
		outputRT.write();
		outputRT.close();
	}

	@Test																												//Metodo com os testes a serem executados
	public void testSequence(){ 
		PrintStream originalPrintStream = System.out;
		try {
			System.out.println("\n\nExecutando casos de testes. Aguarde...\n");
			System.setOut(new PrintStream(workdir + "/Execucao_RT_" + date + ".log")); 		//Salva toda a sa�da do console em um arquivo
			WebServices tc = new WebServices (this.inputRT,this.outputRT, this.url, this.user, this.password);
			String ossStructure = "oss";
			String wsStructure = "ws";

			/* ===================================================================================================================================
			 * INTEGRA��ES WEBSERVICES OSS
			 * ===================================================================================================================================
			*/
			
			System.out.println("\nExecutando casos de teste de Consulta Viabilidade. Aguarde...\n");
 		    tc.testWS(17,ossStructure,"AllocateInstallResource","determineResourceAvailability","<head:context>.*</all:determineResourceAvailabilityRequest>");
	    		   			
			/* ===================================================================================================================================*/ 	
	        
			System.setOut(originalPrintStream);																			//As mensagens serao escritas novamente no console
			System.out.println("O arquivo de log e o roteiro com os dados da execucao foram armazenados dentro do diretorio \"target\".");
			assertTrue(tc.areTestsOk());																				//Gera erro caso algum caso de teste execute falhe
			System.out.println("Execucao finalizada. Todos os casos de teste do roteiro (.xls) executaram com sucesso!\n\n");
		}
		catch (AssertionError error) {
			System.out.println("\nExecucao finalizada com falha! Algum caso de teste do roteiro (.xls) retornou erro! Verifique o roteiro com os dados da execucao para mais detalhes! \n");
			assertTrue(false);
		}
		catch (Exception e) {
			System.setOut(originalPrintStream);																			//As mensagens serao escritas novamente no console
			System.out.println("\nO arquivo de log e o roteiro com os dados da execucao foram armazenados dentro do diretorio \"target\". \n");
			System.out.println ("Finalizado com falha! A ferramenta nao conseguiu os casos de teste. Ocorreu um erro: " + e.toString() + " - verifique os dados de configuracao no RT. \n\n");
		}
	}
}
