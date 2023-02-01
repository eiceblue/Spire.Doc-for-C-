#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"CommentSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"RemoveAndReplaceComment.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Replace the content of the first comment
	doc->GetComments()->GetItem(0)->GetBody()->GetParagraphs()->GetItem(0)->Replace(L"This is the title", L"This comment is changed.", false, false);

	//Remove the second comment
	doc->GetComments()->RemoveAt(1);

	//Save and launch
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
