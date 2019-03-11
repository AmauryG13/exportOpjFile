// exportOpjFile.cpp : Ce fichier contient la fonction 'main'. L'exécution du programme commence et se termine à cet endroit.
//

#include "pch.h"
#include "OriginFile.h"
#include <iostream>

int main(int argc, char *argv[])
{
	if (argc < 2) {
		std::cout << "Usage: exportOpjFile.exe <file.opj>" << endl;
		return -1;
	}

	string input = argv[argc - 1];
	OriginFile opj(input);

	int status = opj.parse();

	std::cout << "Parsing status = " << status << endl;
	std::cout << "OPJ PROJECT \"" << input.c_str() << "\" VERSION = " << opj.version() << endl;

	std::cout << "number of datasets     = " << opj.datasetCount() << endl;
	std::cout << "number of spreadsheets = " << opj.spreadCount() << endl;
	std::cout << "number of matrixes     = " << opj.matrixCount() << endl;
	std::cout << "number of excels       = " << opj.excelCount() << endl;
	std::cout << "number of functions    = " << opj.functionCount() << endl;
	std::cout << "number of graphs       = " << opj.graphCount() << endl;
	std::cout << "number of notes        = " << opj.noteCount() << endl;

}

// Exécuter le programme : Ctrl+F5 ou menu Déboguer > Exécuter sans débogage
// Déboguer le programme : F5 ou menu Déboguer > Démarrer le débogage

// Conseils pour bien démarrer : 
//   1. Utilisez la fenêtre Explorateur de solutions pour ajouter des fichiers et les gérer.
//   2. Utilisez la fenêtre Team Explorer pour vous connecter au contrôle de code source.
//   3. Utilisez la fenêtre Sortie pour voir la sortie de la génération et d'autres messages.
//   4. Utilisez la fenêtre Liste d'erreurs pour voir les erreurs.
//   5. Accédez à Projet > Ajouter un nouvel élément pour créer des fichiers de code, ou à Projet > Ajouter un élément existant pour ajouter des fichiers de code existants au projet.
//   6. Pour rouvrir ce projet plus tard, accédez à Fichier > Ouvrir > Projet et sélectionnez le fichier .sln.
