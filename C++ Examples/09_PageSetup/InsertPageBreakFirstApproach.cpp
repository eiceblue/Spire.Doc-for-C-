#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Template_Docx_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertPageBreakFirstApproach.docx";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str());

	//Find the specified word "technology" where we want to insert the page break.
	vector<TextSelection*> selections = document->FindAllString(L"technology", true, true);

	//Traverse each word "technology".
	for (auto ts : selections)
	{
		TextRange* range = ts->GetAsOneRange();
		Paragraph* paragraph = range->GetOwnerParagraph();
		int index = paragraph->GetChildObjects()->IndexOf(range);

		//Create a new instance of page break and insert a page break after the word "technology".
		Break* pageBreak = new Break(document, BreakType::PageBreak);
		paragraph->GetChildObjects()->Insert(index + 1, pageBreak);
	}

	//Save to file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document->Close();
	delete document;
}
