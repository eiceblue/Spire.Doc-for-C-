#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextBoxTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveTextBox.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Remove the first text box
	doc->GetTextBoxes()->RemoveAt(0);

	//Clear all the text boxes
	//Doc.TextBoxes.Clear();

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
