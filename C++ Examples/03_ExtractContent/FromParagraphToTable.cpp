#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"IncludingTable.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"FromParagraphToTable.docx";

	//Create the first document
	Document* sourceDocument = new Document();

	//Load the source document from disk.
	sourceDocument->LoadFromFile(inputFile.c_str());

	//Create a destination document
	Document* destinationDoc = new Document();

	//Add a section
	Section* destinationSection = destinationDoc->AddSection();

	//Extract the content from the first paragraph to the first table
	ExtractByTable(sourceDocument, destinationDoc, 1, 1);

	//Save the document.
	destinationDoc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	sourceDocument->Close();
	destinationDoc->Close();
	delete sourceDocument;
	delete destinationDoc;
}

void ExtractByTable(Document* sourceDocument, Document* destinationDocument, int startPara, int tableNo)
{
	//Get the table from the source document
	//Table* table = dynamic_cast<Table*>(sourceDocument->GetSections()->GetItem(0)->GetTables()->GetItem(tableNo - 1));
	Table* table = dynamic_cast<Table*>(sourceDocument->GetSections()->GetItem(0)->GetTables()->GetItemInTableCollection(tableNo - 1));

	//Get the table index
	int index = sourceDocument->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->IndexOf(table);
	for (int i = startPara - 1; i <= index; i++)
	{
		//Clone the ChildObjects of source document
		DocumentObject* doobj = sourceDocument->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->GetItem(i)->Clone();

		//Add to destination document 
		destinationDocument->GetSections()->GetItem(0)->GetBody()->GetChildObjects()->Add(doobj);
	}
}