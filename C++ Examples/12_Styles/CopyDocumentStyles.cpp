#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Template_Toc.docx";
	wstring inputFile_2 = input_path  + L"Template_N3.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CopyDocumentStyles.docx";
	
	//Load source document from disk
	Document* srcDoc = new Document();
	srcDoc->LoadFromFile(inputFile.c_str());

	//Load destination document from disk
	Document* destDoc = new Document();
	destDoc->LoadFromFile(inputFile_2.c_str());

	//Get the style collections of source document
	StyleCollection* styles = srcDoc->GetStyles();

	//Add the style to destination document
	for (int i = 0; i < styles->GetCount(); i++)
	{
		IStyle *style = styles->GetItem(i);
		Style* destStyle = dynamic_cast<Style*>(style);
		destDoc->GetStyles()->Add(destStyle);
	}

	//Save the Word file
	destDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	destDoc->Close();
}	
