!INC Local Scripts.EAConstants-JScript

/*
 * Script Name: ListInteractions-JScript
 * Author: Gregory Nikolaidis
 * Purpose: Generate textual descriptions out of a sequence diagram
 * Date: 13-APR-2015
 */
 
function ListInteractions(ObjectID)
{

	var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
	xml = xml + "<EADATA>";
	xml = xml + "<Dataset_0>";
	xml = xml + "<Data>";
	
	var currentDiagram as EA.Diagram;
	currentDiagram = Repository.GetCurrentDiagram;
	
	Session.Output("================================================================");
	Session.Output(" SEQUENCE DIAGRAM - INTERACTION LINSTING");
	Session.Output("----------------------------------------------------------------");
	Session.Output("      Current Diagram ID: " + currentDiagram.DiagramID);
	Session.Output("    Current Diagram Name: " + currentDiagram.Name);

	// ----------------------------------------------------------------
	// Find how many InteractionFragrments for the current sequence
	// diagram
	//
	var query = "SELECT  t_object.ea_guid, t_object.Name, t_diagramobjects.RectTop, t_diagramobjects.RectLeft, " + 
				"t_diagramobjects.RectBottom, t_diagramobjects.RectRight " +
				"FROM t_diagramobjects, t_object " +
				"WHERE t_diagramobjects.Object_ID = t_object.Object_ID " +
				"AND t_object.Object_Type = 'InteractionFragment' " +
				"AND t_diagramobjects.Diagram_ID = " + currentDiagram.DiagramID;
	
	var result = Repository.SQLQuery(query);
	var xmlDOM = XMLParseXML(result);

	var nodePath = "//EADATA/Dataset_0/Data/Row";
	var nodeList = xmlDOM.documentElement.selectNodes(nodePath);
	var numberOfRows = nodeList.length;
	// Session.Output("Number of InteractionFragments: " + numberOfRows);

	// ----------------------------------------------------------------
	// Build JSON with all fragment/partition descriptions and vertical 
	// coordinates
	//
	var have_fragments_to_process = true;

	if (numberOfRows == 0) {
		have_fragments_to_process = false;
	}
	

	if (have_fragments_to_process) {
		var interaction_fragments_json = '[';
		var interaction_fragments = {};
		var partition = Array();

		// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		// Loop through all interactions
		//
		for (var i = 0; i < numberOfRows; i++) {
			
			interaction_fragments_json.concat('{"interaction_fragment": [');
			
			// Session.Output("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - ");
			// Session.Output("             Fragment: " + i);

			var currentRowXPath = "//EADATA/Dataset_0/Data/Row[" + i + "]/";

			var xp_RectTop = currentRowXPath + "RectTop/text()";
			var RectTop = XMLGetNodeText(xmlDOM, xp_RectTop);
			
			var xp_RectBottom = currentRowXPath + "RectBottom/text()";
			var RectBottom = XMLGetNodeText(xmlDOM, xp_RectBottom);
			
			var xp_Name = currentRowXPath + "Name/text()";
			var strName = XMLGetNodeText(xmlDOM, xp_Name);
			
			// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
			// Add the fragment name
			// 
			interaction_fragments_json = interaction_fragments_json + '{"interaction_fragment_name": "' + strName + '", "interaction_fragment": [';
			var xPath_GUID = currentRowXPath + "ea_guid/text()";
			var GUID = XMLGetNodeText(xmlDOM, xPath_GUID);
			var q = "SELECT  t_xref.Description " +
					"FROM t_xref, t_object " +
					"WHERE t_xref.Client = t_object.ea_guid " +
					"AND t_object.ea_guid = '" + GUID + "'";
		
			var r = Repository.SQLQuery(q);
			var dom = XMLParseXML(r);
			var description = XMLGetNodeText(dom, "//EADATA/Dataset_0/Data/Row[0]/Description/text()");

			if (description != "") {
				
				// - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
				// Parse the partitions
				// 
				var partitions = description.split("@ENDPAR;");
				var numOfPartitions = partitions.length - 1;
				// Session.Output(" Number of Partitions: " + numOfPartitions);
				var top = parseInt(RectTop);
				var bottom = top;
				
				for (var j = numOfPartitions; j >= 0; j--) {
					
					if (partitions[j] != "" && partitions[j] != null) {
						
						partition = partitions[j].split(";");
						
						// partition indices 
						//
						// 0: @PAR;	- Don't use
						// 1: Name	- the name of the partition (yes/no, case1/case2/etc.)
						// 2: Size	- the height of the partition
						// 3: GUID	- Don't use
						//
						
						// Get the Name of the partition
						//
						var partition_parts = partition[1].split("=");
						var partition_name = partition_parts[1];
						
						// Get the Name of the partition
						//
						var partition_parts = partition[2].split("=");
						var partition_size = partition_parts[1];
						
						//  -------------------------------------------------
						// Add the partition name and heights (top, bottom)
						// 
						interaction_fragments_json = interaction_fragments_json + '{"name": "' + partition_name + '", ';
						interaction_fragments_json = interaction_fragments_json + '"top": "' + top + '", ';
						bottom = top - parseInt(partition_size);
						interaction_fragments_json = interaction_fragments_json + '"bottom": "' + bottom + '"}';

						top = bottom;
						
						if (j > 0) {
							interaction_fragments_json = interaction_fragments_json + ', ';
						}
						else {
							interaction_fragments_json = interaction_fragments_json + ']';
						}
					}
					else {
						
						// Partition is empty, so skip it
					}
				}

				interaction_fragments_json = interaction_fragments_json + '}';
			}
			else {
				
				//  -------------------------------------------------
				// 
				interaction_fragments_json = interaction_fragments_json + '{"name": "", ';
				interaction_fragments_json = interaction_fragments_json + '"top": "' + RectTop + '", ';
				interaction_fragments_json = interaction_fragments_json + '"bottom": "' + RectBottom + '"}]}';					
			}
			if (i < (numberOfRows - 1)) {

				interaction_fragments_json = interaction_fragments_json + ', ';
			}
		}
		
		interaction_fragments_json = interaction_fragments_json + ']';
		var all_fragments = eval(interaction_fragments_json);
	}
	else {

		// No fragments found in this sequence diagram
		// Session.Output("No Fragments Found in this Sequence Diagram");
	}

	// ----------------------------------------------------------------
	// Get all interactions for the current sequence diagram
	//
	var query = "SELECT t_connector.* " +
				"FROM t_diagram, t_connector " +
				"WHERE t_diagram.Diagram_ID = t_connector.DiagramID " +
				"AND t_diagram.Diagram_ID = " + currentDiagram.DiagramID + " " +
				"ORDER BY t_connector.SeqNo ASC ";
	
	var result = Repository.SQLQuery(query);
	var xmlDOM = XMLParseXML(result);


	// ----------------------------------------------------------------
	// Find how many interactions for the current 
	// sequence diagram
	//
	var nodePath = "//EADATA/Dataset_0/Data/Row";
	var nodeList = xmlDOM.documentElement.selectNodes( nodePath );
	var numberOfRows = nodeList.length;
	// Session.Output("Number of Interactions: " + numberOfRows);

	
	// ----------------------------------------------------------------
	// Initialise some variables
	//
	var majorInteractionNumber = 1;
	var minorInteractionNumber = 0;
	var previous_interaction_fragment_text = "";

	var component_name = new Array();
	
	// ----------------------------------------------------------------
	// Loop through all interactions
	//
	for (var i = 0; i < numberOfRows; i++) {
				
		var currentRowXPath = "//EADATA/Dataset_0/Data/Row[" + i + "]/";
		
		// ----------------------------------------------------------------
		// Check if this interaction is a new activation and 
		// if yes update the major interaction number accordingly
		// 
		var xPath_StateFlags = currentRowXPath + "StateFlags/text()";
		var StateFlags = XMLGetNodeText(xmlDOM, xPath_StateFlags);
		var n = StateFlags.search(/Initiate=1/i);
		if (n >= 0) {
			majorInteractionNumber++;
			minorInteractionNumber = 0;
		}
		var interactionNumber = majorInteractionNumber + "." + minorInteractionNumber;
		




		// ================================================================
		// BUILD THE INTERACTION DESCRIPTION TEXT
		//
		var descriptionText = "";			// This is the final description text to be printed in the doc
		var componentFrom = "";				// This is the source component name
		var componentTo = "";				// This is the target component name
		var noteText = "";					// This is the note text for the current interaction
		var respondsWith = "";				// This is the text part with the response payload
		var interactionDescription = "";	//
		var interactionNotes					// 
		
		// ---------------------------------------------------
		// Get Component Names
		// 
		// COMPONENT FROM --->
		//
		var xPath_Start_Object_ID = currentRowXPath + "Start_Object_ID/text()";
		var Start_Object_ID = XMLGetNodeText(xmlDOM, xPath_Start_Object_ID);
		var query = "SELECT t_object.* FROM t_object WHERE t_object.Object_ID = " + Start_Object_ID;
		var result = Repository.SQLQuery(query);
		var xml_dom_object = XMLParseXML(result);
		componentFrom = XMLGetNodeText(xml_dom_object, "//EADATA/Dataset_0/Data/Row/Name/text()");
		componentFromAlias = XMLGetNodeText(xml_dom_object, "//EADATA/Dataset_0/Data/Row/Alias/text()");

		if (componentFromAlias != "") {
			
			// remove any text in parentheses
			var componentFrom = componentFrom.replace(/\s\(.*\)/i, "");
			
			// Session.Output("component_name["+Start_Object_ID+"] = '" + component_name[Start_Object_ID] + "'" );
			//typeof array[index] !== 'undefined' && array[index] !== null
			if (component_name[Start_Object_ID] == null) { // && component_name[Start_Object_ID] != null && component_name[Start_Object_ID] != "undefined") {
			
				// Session.Output("new");
				componentFrom = componentFrom + " (" + componentFromAlias + ")";
				component_name[Start_Object_ID] = "already seen";
			}
			else {
				componentFrom = componentFromAlias;
			}
		}
		//Session.Output("componentFrom: "+componentFrom);	
		
		
		// ----> COMPONENT TO
		//
		var xPath_End_Object_ID = currentRowXPath + "End_Object_ID/text()";
		var End_Object_ID = XMLGetNodeText(xmlDOM, xPath_End_Object_ID);
		var query = "SELECT t_object.* FROM t_object WHERE t_object.Object_ID = " + End_Object_ID;
		var result = Repository.SQLQuery(query);
		var xml_dom_object = XMLParseXML(result);
		componentTo = XMLGetNodeText(xml_dom_object, "//EADATA/Dataset_0/Data/Row/Name/text()");
		componentToAlias = XMLGetNodeText(xml_dom_object, "//EADATA/Dataset_0/Data/Row/Alias/text()");
		
		if (componentToAlias != "") {

			// remove any text in parentheses
			var componentTo = componentTo.replace(/\s\(.*\)/i, "");
			
			// Session.Output("component_name["+End_Object_ID+"] = '" + component_name[End_Object_ID] + "'" );

			if (component_name[End_Object_ID] == null) { // && component_name[End_Object_ID] != null && component_name[End_Object_ID] != "undefined") {
				
				// Session.Output("new");
				componentTo = componentTo + " (" + componentToAlias + ")";
				component_name[End_Object_ID] = "already seen";
			}
			else {
				componentTo = componentToAlias;
			}
		}
		//Session.Output("componentTo: "+componentTo);	
		
		// ---------------------------------------------------
		// Get the operation name (for call case)
		//
		var xPath_Name = currentRowXPath + "Name/text()";
		var theName = XMLGetNodeText(xmlDOM, xPath_Name);
		
		// ---------------------------------------------------
		// Get operation Notes
		//
		var xPath_Note = currentRowXPath + "Notes/text()";
		interactionNotes = XMLGetNodeText(xmlDOM, xPath_Note);
		
		
		// ---------------------------------------------------
		// Get response description (for response case)
		// 
		var xPath_PDATA2 = currentRowXPath + "PDATA2/text()";
		var PDATA2 = XMLGetNodeText(xmlDOM, xPath_PDATA2);
		var myRegexp = /retval=(.+?);/g;
		var match = myRegexp.exec(PDATA2);
		
		if (match == null) {
			respondsWith = "";
		}
		else if (match[1] != "void") {
			respondsWith = match[1];
		}
		else {
			respondsWith = " empty response";
		}
		
		// ---------------------------------------------------
		// Set the description based on Call/Response 
		// interaction type. The response type is determined 
		// using the value of the PDATA4 column where
		//
		// 0: This is a call
		// 1: This is a resposne
		//
		var xPath_PDATA4 = currentRowXPath + "PDATA4/text()";
		var PDATA4 = XMLGetNodeText(xmlDOM, xPath_PDATA4);
		
		if (PDATA4 == "0") {
			
			if (parseInt(Start_Object_ID) != parseInt(End_Object_ID)) {
				
				// This is a call
				//
				interactionDescription = "" + componentFrom + " calls " + theName + " on " + componentTo;
			}
			else {
				
				// This is a not a call, it is an internal process (could be a call to an internal method)
				//
				if (theName != "") {
					
					// An operation/method is defined for this call
					//
					interactionDescription = "" + componentFrom + " invokes " + theName + " internally to get " + respondsWith;
				}
				else {
					
					// Only a response is defined for this interaction. By convention just pass the text in the response.
					//
					interactionDescription = "" + respondsWith;
				}
			}
		}
		else {
			
			// ---------------------------------------------------
			// This is a response

			var rw = trimAndUpper(respondsWith);

			// If response is empty or response = 'Response' then remove the ' with '
			var response_text = "";
			
			if (rw == "" || rw == "RESPONSE") {
				//do nothing
				response_text = "";
			}
			else {
				response_text = " with " + respondsWith;
			}
			
			interactionDescription = ""  + componentFrom + " responds to " + componentTo + response_text;			
		}

		if (have_fragments_to_process) {
			// ---------------------------------------------------
			// Add additional text to the interaction description 
			// if the current interaction is inside an interaction
			// fragment. 
			//
			// The value to use is PtStartY
			//
			var xp_PtStartY = currentRowXPath + "PtStartY/text()";
			var PtStartY = XMLGetNodeText(xmlDOM, xp_PtStartY);
			
			// - - - - - - - - - - - - - - 
			// Count number of fragments
			//
			var f_index = 0;
			while (all_fragments[f_index]) {
				f_index++;
			}
			var number_of_fragments = f_index;
			// Session.Output("number of fragments: " + number_of_fragments);
			
			var interaction_fragment_text = "";
			var if_and = "";
			var f_index = 0;

			for (var z = 0; z < number_of_fragments; z++) {

				if (z == 0) {
					if_and = "If ";
				}
				else {
					if_and = " and ";
				}
				
				// - - - - - - - - - - - - - - -
				// Count number of partitions
				//
				var p_index = 0;
				while (all_fragments[z].interaction_fragment[p_index]) {
					
					p_index++;
				}
				var number_of_partitions = p_index;
				
				if (number_of_partitions == 1) {
					
					// This is a true/false type of interaction fragment
					//					
					if (parseInt(all_fragments[z].interaction_fragment[0].top) > parseInt(PtStartY) &&
						parseInt(PtStartY) > parseInt(all_fragments[z].interaction_fragment[0].bottom)
						) {
					
						// The current interaction is inside an interaction fragment
						//
						interaction_fragment_text =  interaction_fragment_text + if_and + all_fragments[z].interaction_fragment_name;
					}
					
				}				
				else {

					// This is a case type of interaction fragment
					for (var x = 0; x < number_of_partitions; x++) {
					
						if (parseInt(all_fragments[z].interaction_fragment[x].top) > parseInt(PtStartY) &&
							parseInt(PtStartY) > parseInt(all_fragments[z].interaction_fragment[x].bottom)) {
						
							if (interaction_fragment_text == "") {
								if_and = "If ";
							}
							else {
								if_and = " and ";
							}

							if (number_of_partitions == 2 && (
								all_fragments[z].interaction_fragment[0].name == "YES" || 
								all_fragments[z].interaction_fragment[0].name == "NO" || 
								all_fragments[z].interaction_fragment[0].name == "TRUE" ||
								all_fragments[z].interaction_fragment[0].name == "FALSE"
								) 
							)
							{
								
								// this is a possible true/false yes/no case
								//
								interaction_fragment_text = interaction_fragment_text + if_and + 
															fragmentText(all_fragments[z].interaction_fragment_name, all_fragments[z].interaction_fragment[x].name);
							}
							else {
							
								interaction_fragment_text =  interaction_fragment_text + if_and + 
															all_fragments[z].interaction_fragment_name + " is " + 
															all_fragments[z].interaction_fragment[x].name;
							}
						}				
					}
				}
			}
						
			if (interaction_fragment_text != "" && interaction_fragment_text != previous_interaction_fragment_text) {

				// Previous and current interaction descriptions are different, so use them
				//
				interactionDescription = interaction_fragment_text + " then, " + interactionDescription;
				previous_interaction_fragment_text = interaction_fragment_text;
			}
			else {

				// Previous and current interaction descriptions are the same, so don't use
				//
				interactionDescription = interactionDescription;
			}
		}
		

		// ---------------------------------------------------
		// Add the note
		interactionDescription = interactionDescription + "\r\n" + interactionNotes;

		// ===================================================
		// Build XML results row
		//
		xml = xml + "<Row>";
		xml = xml + "<interactionNumber>" + interactionNumber + "</interactionNumber>";
		xml = xml + "<interactionDescription>" + interactionDescription + "</interactionDescription>";
		xml = xml + "</Row>";

		
		minorInteractionNumber++;		
	}

	xml = xml + "</Data>";
	xml = xml + "</Dataset_0>";
	xml = xml + "</EADATA>";
	// Session.Output(xml);
	return xml;

}


// ==========================================================================================
// ==========================================================================================
// ==========================================================================================
// ==========================================================================================

function trimAndUpper(x) {
	var temp = "";
	temp = x.replace(/^\s+|\s+$/gm,''); // str.trim() does not work!!!
	temp = temp.toUpperCase();
	return temp;
}

// ------------------------------------------------------------------------------------------

function fragmentText(fragment_text, state) {
		
	var updated_fragment_text = fragment_text;
	var s = trimAndUpper(state)
		
	if (s == "YES") {
		// do nothing
	}
	else if (s == "NO") {
		if (fragment_text.search(" is ") >=0) { 
			//Session.Output("Found 'is'");
			updated_fragment_text = fragment_text.replace(" is ", " is not ");
		}
		else if (fragment_text.search(" was ") >=0) { 
			//Session.Output("Found 'was'");
			updated_fragment_text = fragment_text.replace(" was ", " was not ");
		}
		else if (fragment_text.search(" has ") >=0) { 
			//Session.Output("Found 'has'");
			updated_fragment_text = fragment_text.replace(" has ", " doesn't have ");
		}
		else if (fragment_text.search(" have ") >=0) { 
			//Session.Output("Found 'have'");
			updated_fragment_text = fragment_text.replace(" have ", " do not have ");
		}
	}

	//Session.Output("updated_fragment_text: " + updated_fragment_text);
	return updated_fragment_text;
}

// ------------------------------------------------------------------------------------------


// ==========================================================================================
/**
 * Parses a string containing an XML document into an XML DOMDocument object.
 *
 * @param[in] xmlDocument (String) A String value containing an XML document.
 *
 * @return An XML DOMDocument representing the parsed XML Document. If the document could not be 
 * parsed, the function will return null. Parse errors will be logged at the WARNING level
 */
function XMLParseXML( xmlDocument /* : String */ ) /* : MSXML2.DOMDocument */
{
	// Create a new DOM object
	var xmlDOM = new ActiveXObject( "MSXML2.DOMDocument" );
	xmlDOM.validateOnParse = false;
	xmlDOM.async = false;
	
	// Parse the string into the DOM
	var parsed = xmlDOM.loadXML( xmlDocument );
	if ( !parsed )
	{
		// A parse error occured, so log the last error and set the return value to null
		LOGWarning( _XMLDescribeParseError(xmlDOM.parseError) );
		xmlDOM = null;
	}
	
	return xmlDOM;
}

// ------------------------------------------------------------------------------------------
/**
 * Parses an XML file into an XML DOMDocument object.
 *
 * @param[in] xmlPath (String) A String value containing the path name to the XML file to parse.
 *
 * @return An XML DOMDocument representing the parsed XML File.  If the document could not be 
 * parsed, the function will return null. Parse errors will be logged at the WARNING level
 */
function XMLReadXMLFromFile( xmlPath /* : String */ ) /* : MSXML2.DOMDocument */
{
	var xmlDOM = new ActiveXObject( "MSXML2.DOMDocument" );
	xmlDOM.validateOnParse = true;
	xmlDOM.async = true;

	var loaded = xmlDOM.load( xmlPath );
	if ( !loaded )
	{
		LOGWarning( _XMLDescribeParseError(xmlDOM.parseError) );
		xmlDOM = null;
	}
	
	return xmlDOM;
}

// ------------------------------------------------------------------------------------------
/**
 * Saves the provided DOMDocument to the specified file path.
 *
 * @parameter[in] xmlDOM (MSXML2.DOMDocument) The XML DOMDocument to save
 * @parameter[in] filePath (String) The path to save the file to
 * @parameter[in] xmlDeclaration (Boolean) Whether the XML declaration should be included in the 
 * output file
 * @parameter[in] indent (Boolean) Whether the output should be formatted with indents
 */
function XMLSaveXMLToFile( xmlDOM /* : MSXML2.DOMDocument */, filePath /* : String */ , 
	xmlDeclaration /* : Boolean */, indent /* : Boolean */ ) /* : void */
{
	// Create the file to write out to
	var fileIOObject = new ActiveXObject( "Scripting.FileSystemObject" );
	var outFile = fileIOObject.CreateTextFile( filePath, true );
	
	// Create the formatted writer
	var xmlWriter = new ActiveXObject( "MSXML2.MXXMLWriter" );
	xmlWriter.omitXMLDeclaration = !xmlDeclaration;
	xmlWriter.indent = indent;
		
	// Create the sax reader and assign the formatted writer as its content handler
	var xmlReader = new ActiveXObject( "MSXML2.SAXXMLReader" );
	xmlReader.contentHandler = xmlWriter;
	
	// Parse and write the output
	xmlReader.parse( xmlDOM );
	outFile.Write( xmlWriter.output );
	outFile.Close();
}

// ------------------------------------------------------------------------------------------
/**
 * Retrieves the value of the named attribute that belongs to the node at nodePath.
 *
 * @param[in] xmlDOM (MSXML2.DOMDocument) The XML document that the node resides in
 * @param[in] nodePath (String) The XPath path to the node that contains the desired attribute
 * @param[in] attributeName (String) The name of the attribute whose value will be retrieved
 *
 * @return A String representing the value of the requested attribute
 */
function XMLGetAttributeValue( xmlDOM /* : MSXML2.DOMDocument */, nodePath /* : String */, 
	attributeName /* : String */ ) /* : String */
{
	var value = "";
	
	// Get the node at the specified path
	var node = xmlDOM.selectSingleNode( nodePath );
	if ( node )
	{
		// Get the node's attributes
		var attributeMap = node.attributes;
		if ( attributeMap != null )
		{
			// Get the specified attribute
			var attribute = attributeMap.getNamedItem( attributeName )
			if ( attribute != null )
			{
				// Get the attribute's value
				value = attribute.value;
			}
			else
			{
				// Specified attribute not found
				LOGWarning( "Node at path " + nodePath + 
					" does not contain an attribute named: " + attributeName );
			}
		}
		else
		{
			// Node cannot contain attributes
			LOGWarning( "Node at path " + nodePath + " does not contain attributes" );
		}
	}
	else
	{
		// Specified node not found
		LOGWarning( "Node not found at path: " + nodePath );
	}
	
	return value;
}

// ------------------------------------------------------------------------------------------
/**
 * Returns the text value of the XML node at nodePath
 *
 * @param[in] xmlDOM (MSXML2.DOMDocument) The XML document that the node resides in
 * @param[in] nodePath (String) The XPath path to the desired node
 *
 * @return A String representing the desired node's text value
 */
function XMLGetNodeText( xmlDOM /* : MSXML2.DOMDocument */, nodePath /* : String */ ) /* : String */
{
	var value = "";
	
	// Get the node at the specified path
	var node = xmlDOM.selectSingleNode( nodePath );
	if ( node != null )
	{
		value = node.text;
	}
	else
	{
		value = "";
		// Specified node not found
		//LOGWarning( "Node not found at path: " + nodePath );	
	}
	
	return value;
}

// ------------------------------------------------------------------------------------------
/**
 * Returns an array populated with the text values of the XML nodes at nodePath
 *
 * @param[in] xmlDOM (MSXML2.DOMDocument) The XML document that the nodes reside in
 * @param[in] nodePath (String) The XPath path to the desired nodes
 *
 * @return An array of Strings representing the text values of the desired nodes
 */
function XMLGetNodeTextArray( xmlDOM /* : MSXML2.DOMDocument */, nodePath /* : String */ ) /* : Array */
{
	var nodeList = xmlDOM.documentElement.selectNodes( nodePath );
	var textArray = new Array( nodeList.length );
	
	for ( var i = 0 ; i < nodeList.length ; i++ )
	{
		var currentNode = nodeList.item( i );
		textArray[i] = currentNode.text;
	}
	
	return textArray;
}

/**
 * Returns a description of the provided parse error
 *
 * @return A String description of the last parse error that occured
 */
function _XMLDescribeParseError( parseError )
{
	var reason = "Unknown Error";
	
	// If we have an error
	if ( typeof(parseError) != "undefined" )
	{
		// Format a description of the error
		reason = "XML Parse Error at [line: " + parseError.line + ", pos: " + 
			parseError.linepos + "] " + parseError.reason;
	}
	
	return reason;
}


//ListInteractions();