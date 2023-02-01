#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile_1 = input_path + L"Template.docx";
	wstring inputFile_2 = input_path + L"E-iceblue.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ReplaceWithImage.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile_1.c_str());
	//Find the string "E-iceblue" in the document
	vector<TextSelection*> selections = doc->FindAllString(L"E-iceblue", true, true);
	int index = 0;
	TextRange* range = nullptr;

	//Remove the text and replace it with Image
	for (auto selection : selections)
	{
		DocPicture* pic = new DocPicture(doc);
		pic->LoadImageSpire(inputFile_2.c_str());

		range = selection->GetAsOneRange();
		index = range->GetOwnerParagraph()->GetChildObjects()->IndexOf(range);
		range->GetOwnerParagraph()->GetChildObjects()->Insert(index, pic);
		range->GetOwnerParagraph()->GetChildObjects()->Remove(range);
	}

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}