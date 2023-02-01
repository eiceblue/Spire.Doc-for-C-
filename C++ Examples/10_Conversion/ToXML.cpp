#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Summary_of_Science.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToXML.xml";

	//Create word document.
	Document* document = new Document();

	document->LoadFromFile(inputFile.c_str());
	//Save the document to a xml file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Xml);
	document->Close();
	delete document;
}
