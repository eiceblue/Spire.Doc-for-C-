#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Comments.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FromCommentRange.docx";

	//Create a document
	Document* sourceDoc = new Document();

	//Load the document from disk.
	sourceDoc->LoadFromFile(inputFile.c_str());

	//Create a destination document
	Document* destinationDoc = new Document();

	//Add section for destination document
	Section* destinationSec = destinationDoc->AddSection();

	//Get the first comment
	Comment* comment = sourceDoc->GetComments()->GetItem(0);

	//Get the paragraph of obtained comment
	Paragraph* para = comment->GetOwnerParagraph();

	//Get index of the CommentMarkStart 
	int startIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkStart());

	//Get index of the CommentMarkEnd
	int endIndex = para->GetChildObjects()->IndexOf(comment->GetCommentMarkEnd());

	//Traverse paragraph ChildObjects
	for (int i = startIndex; i <= endIndex; i++)
	{
		//Clone the ChildObjects of source document
		DocumentObject* doobj = para->GetChildObjects()->GetItem(i)->Clone();

		//Add to destination document 
		destinationSec->AddParagraph()->GetChildObjects()->Add(doobj);
	}
	//Save the destination document
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	sourceDoc->Close();
	destinationDoc->Close();
	delete sourceDoc;
	delete destinationDoc;
}