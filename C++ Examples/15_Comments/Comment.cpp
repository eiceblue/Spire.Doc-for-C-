#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"CommentTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Comment.docx";

	//Load the document from disk.
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	InsertComments(document->GetSections()->GetItem(0));

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void InsertComments(Section* section) {
	//Insert comment.
	Paragraph* paragraph = section->GetParagraphs()->GetItem(1);
	Spire::Doc::Comment* comment = paragraph->AppendComment(L"Spire.Doc for C++");
	comment->GetFormat()->SetAuthor(L"E-iceblue");
	comment->GetFormat()->SetInitial(L"CM");
}	
