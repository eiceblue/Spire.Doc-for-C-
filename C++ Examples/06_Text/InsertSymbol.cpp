#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertSymbol.docx";

	//Create Word document.
	Document* document = new Document();

	//Add a section.
	Section* section = document->AddSection();

	//Add a paragraph.
	Paragraph* paragraph = section->AddParagraph();

	//Use unicode characters to create symbol Ä.
	wstring tempA = L"\u00c4";
	TextRange* tr = paragraph->AppendText(tempA.c_str());

	//Set the color of symbol Ä.
	tr->GetCharacterFormat()->SetTextColor(Color::GetRed());

	//Add symbol Ë.
	wstring tempB = L"\u00cb";
	paragraph->AppendText(tempB.c_str());

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
