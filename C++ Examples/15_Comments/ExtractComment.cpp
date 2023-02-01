#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"CommentSample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ExtractComment.txt";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	wstring* SB = new wstring();

	//Traverse all comments
	for (int i = 0; i < doc->GetComments()->GetCount(); i++)
	{
		Comment* comment = doc->GetComments()->GetItem(i);
		for (int j = 0; j < comment->GetBody()->GetParagraphs()->GetCount(); j++)
		{
			Paragraph* p = comment->GetBody()->GetParagraphs()->GetItem(j);
			SB->append(p->GetText());
			SB->append(L"\n");
		}
	}

	//Save to TXT File and launch it
	wofstream write(outputFile);
	write << SB->c_str();
	write.close();
	doc->Close();
	delete doc;
	delete SB;
}
