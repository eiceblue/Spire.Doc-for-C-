#include "pch.h"

using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"CommentTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddCommentForSpecificText.docx";

	//Load the document from disk.
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	InsertComments(document, L"development");

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void InsertComments(Document* doc, const wstring& keystring)
{
	//Find the key string
	TextSelection* find = doc->FindString(keystring.c_str(), false, true);

	//Create the commentmarkStart and commentmarkEnd
	CommentMark* commentmarkStart = new CommentMark(doc);
	commentmarkStart->SetType(CommentMarkType::CommentStart);
	CommentMark* commentmarkEnd = new CommentMark(doc);
	commentmarkEnd->SetType(CommentMarkType::CommentEnd);

	//Add the content for comment
	Comment* comment = new Comment(doc);
	comment->GetBody()->AddParagraph()->SetText(L"Test comments");
	comment->GetFormat()->SetAuthor(L"E-iceblue");

	//Get the textRange
	TextRange* range = find->GetAsOneRange();

	//Get its paragraph
	Paragraph* para = range->GetOwnerParagraph();

	//Get the index of textRange 
	int index = para->GetChildObjects()->IndexOf(range);

	//Add comment
	para->GetChildObjects()->Add(comment);

	//Insert the commentmarkStart and commentmarkEnd
	para->GetChildObjects()->Insert(index, commentmarkStart);
	para->GetChildObjects()->Insert(index + 2, commentmarkEnd);
}
