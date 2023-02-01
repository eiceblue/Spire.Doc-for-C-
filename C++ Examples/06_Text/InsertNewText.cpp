#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertNewText.docx";

	//Load Document   
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Find all the text string "New Zealand” from the sample document
	vector<TextSelection*> selections = doc->FindAllString(L"Word", true, true);
	int index = 0;

	//Defines text range
	TextRange* range = new TextRange(doc);

	//Insert new text string (New) after the searched text string
	for (auto selection : selections)
	{
		range = selection->GetAsOneRange();
		TextRange* newrange = new TextRange(doc);
		newrange->SetText(L"(New text)");
		index = range->GetOwnerParagraph()->GetChildObjects()->IndexOf(range);
		range->GetOwnerParagraph()->GetChildObjects()->Insert(index + 1, newrange);
	}

	//Find and highlight the newly added text string New
	vector<TextSelection*> text = doc->FindAllString(L"New text", true, true);
	for (auto seletion : text)
	{
		seletion->GetAsOneRange()->GetCharacterFormat()->SetHighlightColor(Color::GetYellow());
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}