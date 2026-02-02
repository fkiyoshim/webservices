package br.com.cpqd;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.*;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.ws.security.WSConstants;
import org.apache.ws.security.message.WSSecHeader;
import org.apache.ws.security.message.WSSecUsernameToken;

import com.eviware.soapui.SoapUI;
import com.eviware.soapui.impl.WsdlInterfaceFactory;
import com.eviware.soapui.impl.wsdl.WsdlInterface;
import com.eviware.soapui.impl.wsdl.WsdlOperation;
import com.eviware.soapui.impl.wsdl.WsdlProject;
import com.eviware.soapui.impl.wsdl.WsdlSubmit;
import com.eviware.soapui.impl.wsdl.WsdlSubmitContext;
import com.eviware.soapui.impl.wsdl.WsdlRequest;
import com.eviware.soapui.model.iface.Response;
import com.eviware.soapui.settings.HttpSettings;
import com.eviware.soapui.support.xml.XmlUtils;

import org.custommonkey.xmlunit.DetailedDiff;
import org.custommonkey.xmlunit.Diff;

import java.util.List;

public class WebServices {
    private Workbook inputRT;				
    private WritableWorkbook outputRT; 		
    private Sheet inputSheet;				 
    private WritableSheet outputSheet; 	
    private Cell cell[];	
    private String user;
    private String password;
    private WsdlOperation operation = null;
    private boolean testsOk = true;
    private String url;
    
    public WebServices (Workbook inputRT, WritableWorkbook outputRT, String url, String user, String password) throws IOException , Exception {
    	this.inputRT=inputRT;
		this.outputRT=outputRT;
		this.user=user;
		this.password=password;
		this.url=url;
    }
    
    // Transforma uma string em um documento XML
   	private String transformXML (String str) throws Exception {
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();  
		DocumentBuilder builder;  
		builder = factory.newDocumentBuilder();  
		Document document = builder.parse(new InputSource(new StringReader(str)));  
		TransformerFactory tf = TransformerFactory.newInstance();
		Transformer transformer = tf.newTransformer();
		transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");
		StreamResult res = new StreamResult(new StringWriter());
		DOMSource source = new DOMSource(document);
		transformer.transform(source, res);
		return res.getWriter().toString();
   	}
   	
   	// Adiciona o tipo de autenticacao "wss" na requisicao
   	private void addWSSAuthentication (WsdlRequest request){
		if( ( request.getUsername() == null || request.getUsername().length() == 0 ) || ( request.getPassword() == null || request.getPassword().length() == 0 ) )
			System.out.println("ERRO: Informe o Usuario e/ou Senha" );		
		String req = request.getRequestContent();
		try
		{
			String passwordType = WsdlRequest.PW_TYPE_DIGEST;
			if( passwordType == null )
				return;
			DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
			dbf.setNamespaceAware( true );
			DocumentBuilder db = dbf.newDocumentBuilder();
			Document doc = db.parse( new InputSource( new StringReader( req ) ) );
			WSSecUsernameToken addUsernameToken = new WSSecUsernameToken();
			if( WsdlRequest.PW_TYPE_DIGEST.equals( passwordType ) )
				addUsernameToken.setPasswordType( WSConstants.PASSWORD_DIGEST );
			else
				addUsernameToken.setPasswordType( WSConstants.PASSWORD_TEXT );
			addUsernameToken.setUserInfo( request.getUsername(), request.getPassword() );
			addUsernameToken.addNonce();
			addUsernameToken.addCreated();
			StringWriter writer = new StringWriter();
			WSSecHeader secHeader = new WSSecHeader();
			secHeader.insertSecurityHeader( doc );
			XmlUtils.serializePretty( addUsernameToken.build( doc, secHeader ), writer );
			request.setRequestContent( writer.toString() );
		}
		catch( Exception e1 )
		{
			System.out.println("ERRO: Ocorreu um erro na adicao do token de seguranca: " + e1 );
		}
	}
		
	// Compara os dados do resultado esperado x obtido
	private void resultVerify (String structureFolder, String fluxo, String expectedMessage, String obtainedMessage){
		Pattern pat = null;
		Pattern patError = Pattern.compile("<soap:Fault>.*</soap:Envelope>");
		Matcher mat = null;
		Matcher mat2 = null;

		if (structureFolder.equals("oss") || structureFolder.equals("ws"))
			pat = Pattern.compile("result>.*</soap:Envelope>");
				
		mat = pat.matcher(expectedMessage);
		if (mat.find()) {     
			System.out.println("\nExpectedMessage: " + mat.group());
			mat2 = pat.matcher(obtainedMessage);
			if (mat2.find()) {  
				System.out.println("\nObtainedMessage: " + mat2.group());
				if(!mat.group().equals(mat2.group())){
					resultVerifyErrorFault(patError, expectedMessage, obtainedMessage);
				}
			}
			else {
				resultVerifyErrorFault(patError, expectedMessage, obtainedMessage);
			}
		}
		else {
			resultVerifyErrorFault(patError, expectedMessage, obtainedMessage);
		}
	}

	// Compara os dados dos resultados esperados 1 e 2 x obtido - utilizado somente nos casos de consulta viabilidade quando um endereco tem viabilidade/cobertura em redes Fttx e Metalica.
	private void resultVerify (String expectedMessage, String expectedMessage2, String obtainedMessage) {
		Pattern pat = null;
		Pattern patError = Pattern.compile("<soap:Fault>.*</soap:Envelope>");
		Matcher mat = null;
		Matcher mat2 = null;
		Matcher mat3 = null;
	
		pat = Pattern.compile("result>.*</soap:Envelope>");
		
		mat = pat.matcher(expectedMessage);
		if (mat.find()) {     
			System.out.println("\nExpectedMessage 1: " + mat.group());
			mat3 = pat.matcher(expectedMessage2);
			mat3.find();
           	System.out.println("\nExpectedMessage 2: " + mat3.group());
			mat2 = pat.matcher(obtainedMessage);
			if (mat2.find()) {  
				System.out.println("\nObtainedMessage: " + mat2.group());
				if (mat.group().equalsIgnoreCase(mat2.group()) || mat3.group().equalsIgnoreCase(mat2.group())) { 		// Compara o valor da tag indicada nas mensagens esperada 1/2 e obtida
					System.out.println("\n OK!");
				}
				else {																									// Gera uma excecao indicando erro caso o valor da tag indicada seja diferente nas mensagens esperada e obtida
					resultVerifyErrorFault(patError, expectedMessage, obtainedMessage);
				}
			}
			else {
				resultVerifyErrorFault(patError, expectedMessage, obtainedMessage);																		// Gera uma excecao indicando erro caso o valor da tag indicada seja diferente nas mensagens esperada e obtida
			}
		}
		else {
			mat2 = pat.matcher(obtainedMessage);
			mat3 = pat.matcher(expectedMessage2);
			if (mat2.find() ||  mat3.find()) {
            	System.out.println("Nao OK - encontrada somente no resultado obtido!");
            	throw new FailedException();
			}
		}
	}	
	
	//Método que verifica se a mensagem retornada é um erro não estruturado do sistema.
	private void resultVerifyErrorFault(Pattern patError, String expectedMessage, String obtainedMessage) {
			Matcher mat = null;
			Matcher mat2 = null;
			
				mat = patError.matcher(expectedMessage);
				if (mat.find()) {     
					System.out.println("\nExpectedMessage: " + mat.group());
					mat2 = patError.matcher(obtainedMessage);
					if (mat2.find()) {  
						System.out.println("\nObtainedMessage: " + mat2.group());
						if(!mat.group().equals(mat2.group())){
							throw new FailedException();
						}
					}
					else {
						throw new FailedException();
					}
				}
				else {
					throw new FailedException();
				}
	}
	
	/*
		O metodo replaceId realiza o tratamento na mensagem devido a:
			a) NS gerados com ID diferentes -> Serao substituidos por nsX
			b) Campos com valores de IDs -> Conteudos serao substituidos por XXXX
	*/
	private String replaceID (String message) {
		message=message.replaceAll("ns\\d*:", "nsX:");
		message=message.replaceAll("ns\\d*=", "nsX=");
		message=message.replaceAll("\r", ""); 
		message=message.replaceAll("\t", "");
		message=message.replaceAll("\n", "");
		message=message.replaceAll("<nsX:resourceIdentifier>\\d*</nsX:resourceIdentifier>", "<nsX:resourceIdentifier>XXXX</nsX:resourceIdentifier>");
		message=message.replaceAll("REFERENCE\">\\d*","REFERENCE\">XXXX");
		message=message.replaceAll("EQPT_TERM_ID\">\\d*","EQPT_TERM_ID\">XXXX");
		message=message.replaceAll("TERMINAL_BOX_FNUM\">\\d*","TERMINAL_BOX_FNUM\">XXXX");
		message=message.replaceAll("TERMINAL_BOX_ID\">\\d*","TERMINAL_BOX_ID\">XXXX");
		message=message.replaceAll("<nsX:restrictionObjectId>\\d*</nsX:restrictionObjectId>", "<nsX:restrictionObjectId>XXXX</nsX:restrictionObjectId>");
		message=message.replaceAll("<nsX:resourceOrderDate>\\d*-\\d*-\\d*</nsX:resourceOrderDate>", "<nsX:resourceOrderDate>XXXX</nsX:resourceOrderDate>");
		return message;
	}
		
	// Escreve dados/mensagens no roteiro de teste
    @SuppressWarnings("deprecation")
	private void writeResult (Colour background, jxl.format.Alignment alignment, jxl.format.VerticalAlignment verticalAlignment,  int resultColumn, int resultLine, String message) throws Exception {
		Label res;
    	WritableCellFormat wcf = new WritableCellFormat ();
		WritableFont wr = new WritableFont(WritableFont.ARIAL,8);
		wcf.setFont(wr);
		wcf.setBackground(background);
		wcf.setAlignment(alignment);
		wcf.setVerticalAlignment(verticalAlignment);
		wcf.setBorder(Border.ALL, BorderLineStyle.THIN, Colour.WHITE);
		wcf.setWrap(true);
		if (message.length()>32767) {
			res = new Label(resultColumn, resultLine, "O texto possui uma quantidade de caracteres maior do que o limite da celula. Veja o texto original no log!");
			System.out.println("\nTamanho da mensagem (32767 maximo aceitado)=" + message.length());
		}
		else
			res = new Label(resultColumn, resultLine, message);
		res.setCellFormat(wcf);
		outputSheet.addCell(res);
    }
    
	// Executa os testes dos fluxos - os dados de entrada sao obtidos do roteiro de testes
    @SuppressWarnings({ "deprecation", "rawtypes" })
	public void testWS (int sheetTest, String structureFolder, String wsdl, String fluxo, String replaceTag) throws Exception {
    	/*
		  sheetTest indica a planilha com os casos de teste - valores validos:
				1 - reserva
				2 - cancelamento de reserva
				3 - alocacao
				4 - cancelamento de alocacao
				5 - associacao de servicos
				6 - desassociacao de servicos
				7 - consulta recursos
				8 - cadastro de retirada
    			9 - cancelamento de retirada
    		   10 - confirmacao de retirada
    		   11 - consulta recursos em equipamento terminal
    		   12 - manobra forcada de recursos
    		   13 - consulta viabilidade
    		   14 - alteracao do endereco cadastrado na facilidade

    	  wsdl - valores validos:
    			AllocateInstallResource
    			TrackManageResourceProvisioning
    			TrackManageServiceProvisioning
    			ReportResourceProvisioning
    			
    	  fluxo - indica a qual operacao a ser testada - valores validos:
    			reserveResource
    			manageResourceProvisioningActivity
    			allocateResource
    			releaseResource
    			manageServiceProvisioningActivity
    			distributeResourceProvisioningReports
    			distributeFreeResourceTerminalEquipment
    			releaseAllocateResource
    			determineResourceAvailability
    			editCircuitAddress
    			
		  replaceTag - indica a expressao que sera pesquisada no xml de entrada de entrada para ter o conteúdo substituido pelo definido no RT

		*/
    	
    	String requestRT;
		String inMessage=null;
		String expectedMessage=null;
		String expectedMessage2=null;
		String obtainedMessage=null;
		String URLW;
		int resultColumn=0;
		int obtainedMessageColumn=0;
		int expectedMessageColumn=4;
		int inMessageModifiedColumn=3;
		int i=0;
		String expectedMessageOri=null;
		String obtainedMessageOri=null;
		String expectedMessageOri2=null;

		inputSheet = inputRT.getSheet(sheetTest); 	
	    WsdlProject project = new WsdlProject();
	    
	    if(structureFolder != "ws"){
	    	URLW = this.url + structureFolder.toLowerCase() + "/" + wsdl+"?wsdl";
	    }else{
	    	URLW = this.url + "/" + wsdl+"?wsdl";
	    }
	    
	    System.out.println("\n\nWSDL: " + URLW + "\n\n");
	    WsdlInterface iface;
		iface = WsdlInterfaceFactory.importWsdl(project, URLW, false )[0];
		WsdlRequest request=null;
		SoapUI.getSettings().setString(HttpSettings.SOCKET_TIMEOUT,"600000");

		if (fluxo=="determineResourceAvailability") {
			resultColumn=7;
		    obtainedMessageColumn=6;
		}
		else {
			resultColumn=6;
		    obtainedMessageColumn=5;
		}
		
		for (i=5; i < inputSheet.getRows(); i++) {
			try {	
				System.out.println("\n-----------------------------------------------------------------------------------------\n");
				System.out.println("Fluxo: " + fluxo + "\n");
				System.out.println("Caso de teste: " + (i-4) + "\n");
				
				cell=inputSheet.getRow(i);
				requestRT=cell[2].getContents();	
				operation = (WsdlOperation) iface.getOperationByName(fluxo);
				
				// Cria uma nova requisicao
				request = operation.addNewRequest( "Request" );
				
				// Gera o conteudo da requisicao
				request.setRequestContent( operation.createRequest( true ) );
				
				//Adiciona a autenticacao WSS Autentication. Exceto para os fluxos de Consulta Recursos e Consulta Recursos em Equipamento Terminal que utilizam HTTP Basic Autentication.
				request.setUsername(user);
				request.setPassword(password);
				this.addWSSAuthentication(request);
				
				inMessage = request.getRequestContent();
				inMessage = inMessage.replaceAll("\r", ""); 
				inMessage = inMessage.replaceAll("\t", "");
				inMessage = inMessage.replaceAll("\n", "");

				// Substitui o conteudo da variavel replaceTag pela requisicao definida no RT
				inMessage=inMessage.replaceAll(replaceTag,requestRT);

				// Transforma a inMessagem em XML
				inMessage=transformXML(inMessage);
				
				System.out.println("Messagem de entrada: \n" + inMessage + "\n");
				request.setTimeout("900000");   //15min
				
				// Atribui conteudo da variavel inMessage a requisicao a ser testada
				request.setRequestContent(inMessage);
			
				// Execucao da requisicao modificada
				WsdlSubmit submit = (WsdlSubmit) request.submit( new WsdlSubmitContext(request), false );
				submit.getRequest();
				Response response = submit.getResponse();
				obtainedMessage = response.getContentAsString();
				System.out.println("\nMessagem de resposta obtida= \n" + obtainedMessage + "\n");
				
				//obtainedMessage=outMessage;
				outputSheet=outputRT.getSheet(sheetTest);
				
				//Escrita da mensagem de entrada completa no RT
				writeResult (Colour.GRAY_25, Alignment.LEFT, VerticalAlignment.TOP,  inMessageModifiedColumn, i, inMessage);
				
				//Escrita da mensagem de resposta obtida no RT
				writeResult (Colour.GRAY_25, Alignment.LEFT, VerticalAlignment.TOP,  obtainedMessageColumn, i, obtainedMessage);
				
				expectedMessage=cell[expectedMessageColumn].getContents();
				System.out.println("Mensagem de resposta esperada: \n " + expectedMessage);
				if (expectedMessage.isEmpty()){
	            	System.out.println("\nMensagem esperada vazia!");
	            	throw new FailedException();
				}
				
				expectedMessageOri=expectedMessage;
				obtainedMessageOri=obtainedMessage;
				expectedMessage=this.replaceID(expectedMessage);
				obtainedMessage=this.replaceID(obtainedMessage);
					
				if (fluxo=="determineResourceAvailability") {				//Fluxo de consulta viabilidade
					// Verificacao utilizada nos casos que a consulta viabilidade retorna atendimento por redes Metalica e FTTx
					if (expectedMessage.contains("netType>1") && expectedMessage.contains("netType>2")) {   
						System.out.println("\nEndereco atendido por redes metalica e FTTx!");
						expectedMessage2=cell[expectedMessageColumn+1].getContents();
						expectedMessageOri2=expectedMessage2;
						expectedMessage2=this.replaceID(expectedMessage2);
						System.out.println("\nComparacao das mensagens Esperada x Obtida:");
						this.resultVerify(expectedMessage, expectedMessage2, obtainedMessage);
					}
					else {
						System.out.println("\nComparacao das mensagens Esperada x Obtida:");
						this.resultVerify(structureFolder, fluxo, expectedMessage, obtainedMessage);
					}
				} 
				else {
					System.out.println("\nComparacao das mensagens Esperada x Obtida:");
					this.resultVerify(structureFolder, fluxo, expectedMessage, obtainedMessage);
				}
				writeResult (Colour.LIGHT_GREEN, Alignment.CENTRE, VerticalAlignment.CENTRE,  resultColumn, i, "OK");
				System.out.println("OK!");
				writeResult (Colour.GRAY_25, Alignment.LEFT, VerticalAlignment.TOP,  resultColumn+1, i, " ");
			}	
			
			catch (FailedException e) {
				//Em caso de erro no teste serao informadas/indicadas as principais diferencas nos XMLs de resposta esperada x obtida, sendo utilizado xmlunit para verificar essas diferencas
				try {
					
					writeResult (Colour.RED, Alignment.CENTRE, VerticalAlignment.CENTRE,  resultColumn, i, "NOK");
					System.out.println("\nTeste NOK! Falha na comparacao dos XMLs de resposta esperado x obtido");
					Diff diff=null;
					if (fluxo=="determineResourceAvailability") {					//Fluxo de consulta viabilidade
						// Verificacao utilizada nos casos que a consulta viabilidade retorna atendimento por redes Metalica e FTTx
						if (expectedMessageOri.contains("netType>1") && expectedMessageOri.contains("netType>2")) {			
							System.out.println("netType 1 e 2 encontrados! - expectedMessageOri");
							if (obtainedMessageOri.contains("netType>1") && obtainedMessageOri.contains("netType>2")) {
								System.out.println("netType 1 e 2 encontrados! - obtainedMessageOri");
								String auxiliar=obtainedMessageOri;
								auxiliar=auxiliar.replaceAll("\r", "");		
								auxiliar=auxiliar.replaceAll("\t", "");
								auxiliar=auxiliar.replaceAll("\n", "");
								if (auxiliar.matches(".*netType>1.*netType>2.*")) {
									System.out.println("netType1 antes de netType2");
									diff = new Diff(expectedMessageOri, obtainedMessageOri);
								}
								else {
									if (auxiliar.matches(".*netType>2.*netType>1.*")) {
										System.out.println("netType2 antes de netType1");
										diff = new Diff(expectedMessageOri2, obtainedMessageOri);
									}
								}
							}	
						}
						else
							diff = new Diff(expectedMessageOri, obtainedMessageOri);
					}
					else
						diff = new Diff(expectedMessageOri, obtainedMessageOri);
					
					DetailedDiff detailXmlDiff = new DetailedDiff(diff);
	      	        List differences = detailXmlDiff.getAllDifferences();
	      	        
	      	        System.out.println("\nMensagens de erro originais (xmlunit):");  
	      	        for (int j=0;j<differences.size();j++)
	      	        	System.out.println(differences.get(j));
	      	    	 
	      	        System.out.println("\nMensagens de erro filtradas/selecionadas:");  
	      	        String message=null;
	      	        String filterMessage=null;
	      	        int id=1; 
	      	        for(int j=0;j<differences.size();j++){
	      	        	message=differences.get(j).toString();
	      	        	message=message.replaceAll("\r", "");		
	      	        	message=message.replaceAll("\t", "");
	      	        	message=message.replaceAll("\n", "");
	      	        	
	      	        	// Filtro das mensagens de erro mais significativas para o teste
	      	        	if ((!message.matches("Expected namespace prefix.*"))										&& 
	      	        		(!message.matches("Expected number of child.*"))										&& 
	     	        		(!message.matches("Expected sequence.*"))												&&
	      	        		(!message.matches("Expected presence of child node 'null' but was '#text.*"))			&& 
	      	        		(!message.matches("Expected presence of child node '#text' but was 'null.*"))			&&
	      	        		(!message.matches(".*resourceIdentifier.*"))											&&
	      	        		(!message.matches(".*REFERENCE.*"))														&&
	      	        		(!message.matches(".*EQPT_TERM_ID.*"))													&&
	      	        		(!message.matches(".*TERMINAL_BOX_FNUM.*"))												&&
	      	        		(!message.matches(".*TERMINAL_BOX_ID.*"))												&&
	      	        		(!message.matches(".*restrictionObjectId.*"))											&&
	      	        		(!message.matches(".*resourceOrderDate.*"))												&&
	      	        		(!message.matches("Expected text value '[0-9]{7,10}' but was '[0-9]{7,10}'.*"))			&&
	      	        		(!message.matches(".*serviceTypeDesc.*"))												&&
	      	        		(!message.matches(".*resourceOrderDate.*"))												&&
	      	        		(!message.matches("Expected text value '\\s\\s\\s.*"))) {
	      	        			if (filterMessage==null)
	      	        				filterMessage=id + ") " + message + "\n";
	      	        			else
	      	        				filterMessage=filterMessage.concat(id + ") " + message + "\n");
	      	        			id++;
		      	    	 }
	         	    }
	      	        
	      	        if (filterMessage!=null) {
	      	        	// Verificacao se a mensagem nao ultrapassa o limite de caracteres aceito em uma celula do arquivo xls
		      			if (filterMessage.length()>32767) {
		      				System.out.println("\nTamanho da mensagem (32767 maximo aceitado)=" + filterMessage.length());
		      				System.out.println("O texto original possui uma quantidade de caracteres maior do que o limite da celula, por isso foi truncado! Veja o texto original no log!");
		      				String subfilterMessage=filterMessage.substring(0, 32500);
		      				subfilterMessage=subfilterMessage.concat("\nO texto original possui uma quantidade de caracteres maior do que o limite da celula, por isso foi truncado! Veja o texto original no log!\n");
		      				System.out.println(subfilterMessage);
		      				writeResult (Colour.GRAY_25, Alignment.LEFT, VerticalAlignment.TOP,  resultColumn+1, i, subfilterMessage);
		      			}
		      			else {
		          	        System.out.println(filterMessage);
		    				writeResult (Colour.GRAY_25, Alignment.LEFT, VerticalAlignment.TOP,  resultColumn+1, i, filterMessage);
		      			}
	      	        }
					testsOk = false;
				}

				catch (Exception ex) {
					System.out.println("Teste NOK! - Falha na comparacao via xmlunit!");
					System.out.println("Excecao= " + ex.getMessage());
					testsOk = false;
				}
			}
			catch (SAXException e) {
				writeResult (Colour.RED, Alignment.CENTRE, VerticalAlignment.CENTRE,  resultColumn, i, "NOK");
				System.out.println("\nTeste NOK!");
				testsOk = false;
			}
		} 
    }
     
    // Retorna true somente quando todos os casos do roteiro foram executados com sucesso
    public boolean areTestsOk() {
    	return testsOk;
    }
    
}
