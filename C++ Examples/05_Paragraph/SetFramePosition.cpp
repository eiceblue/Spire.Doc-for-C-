#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"TextInFrame.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetFramePosition.docx";

	//Create a document
	Document* document = new Document();

	//Load the document from disk
	document->LoadFromFile(inputFile.c_str());

	//Get a paragraph
	Paragraph* paragraph = document->GetSections()->GetItem(0)->GetParagraphs()->GetItem(0);

	//Set the Frame's position
	if (paragraph->GetFormat()->GetIsFrame())
	{
		paragraph->GetFormat()->GetFrame()->SetHorizontalPosition(150.0f);
		paragraph->GetFormat()->GetFrame()->SetVerticalPosition(150.0f);
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
