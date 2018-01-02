#!/usr/bin/env python3
import argparse
import lxml.etree
import os
import re
import sys
import zipfile


class OfficeMetadata:
    """
    Retrieve MS office metadata, as well as images.
    """
    def __init__(self, item, mediaDir, dumpMedia):
        self.item = item
        self.mediaDir = mediaDir
        self.dm = dumpMedia
        self.xtractor()

    def isPKFile(self, file):
        """
        Check if it's a valid MS office file.
        :param file: file path.
        :return: True if it's valid, if not, blow your PC.
        """
        with open(file, "rb") as fileMagicNumber:
            if fileMagicNumber.read(2) != b"PK":
                return True

        return False

    def hasMediaData(self, compFile):
        """
        Checks if the file has embedded media.
        :param compFile: compressed file path.
        :return: list of the compressed media path.
        """
        regex = re.compile("word/media/([A-Z]|[0-9])*\.(jpeg|gif|png)", re.I)
        mediaFiles = list()

        for file in compFile.namelist():
            if re.match(regex, file):
                mediaFiles.append(file)

        return mediaFiles

    def printData(self, foundData):
        """
        Prints found data.
        :param foundData: Gathered elements.
        """
        for key, value in foundData.items():
            print("\t[-] {0}: {1}".format(key.title(), value))

    def parseXML(self, xmlFile):
        """
        Parses OpenXML compressed file.
        :param xmlFile: compressed filename.
        :return: dictionary containing XML file information.
        """
        print("\n[+] Parsing {}.".format(xmlFile.name))

        xmlElementTree = lxml.etree.parse(xmlFile)
        foundData = dict()

        for element in xmlElementTree.iter("*"):
            tagName = lxml.etree.QName(element.tag).localname
            if element.text is not None:
                foundData[tagName] = element.text

        self.printData(foundData)

    def getMedia(self, extFiles, file, filename):
        """
        Uncompress found media files.
        :param extFiles: media found.
        :param file: zipfile Object.
        :param filename: MS office filename.
        :return:
        """
        path = ""

        if self.mediaDir == ".":
            path = "{0}/{1}".format(self.mediaDir, os.path.splitext(filename)[0])
        else:
            path = self.mediaDir

        relpath = os.path.relpath(path)

        if os.path.exists(relpath):
            print("\n[!] Directory {} exists.".format(relpath))

        else:
            print("[*] Creating directory {}".format(relpath))
            os.mkdir(relpath)

        print("[*] Extracting {0} files in {1}".format(len(extFiles), relpath))
        file.extractall(members=extFiles, path=relpath)

    def getMetadata(self, file):
        """
        Retrieves MS Office file metadata, parsing core.xml file and app.xml file.
        :param file: MS Office file path.
        """
        docxData = None
        try:
            docxData = zipfile.ZipFile(file)

        except zipfile.BadZipfile as zipErr:
            print("{} is not a valid zip file.".format(file))

            if self.isPKFile(self.item):
                print("{} is not a valid office file.".format(file))

            sys.exit()

        print("\n\n[*] Parsing metadata from {}".format(file))

        with docxData.open("docProps/core.xml") as core:
            self.parseXML(core)

        with docxData.open("docProps/app.xml") as app:
            self.parseXML(app)

        mediaFiles = self.hasMediaData(docxData)

        if len(mediaFiles) < 1:
            print("[!] No media detected.")

        else:
            print("[*] {} media file(s) detected.".format(len(mediaFiles)))
            if self.dm:
                self.getMedia(mediaFiles, docxData, file)

    def recursiveSearch(self):
        """
        If specified, the script will recursively parses an entire directory that contains MS Office files.
        """
        regexPattern = re.compile("([A-Z]|[0-9])*\.docx", re.I)
        print("[*] Parsing directory: {}".format(self.item))

        for file in os.listdir(self.item):
            if re.match(regexPattern, file):
                self.getMetadata("{0}/{1}".format(self.item, file))

    def xtractor(self):
        """
        Checks if provided pathname it's a file or a directory.
        :return:
        """
        if os.path.isdir(self.item):
            self.recursiveSearch()
        elif os.path.isfile(self.item):
            self.getMetadata(self.item)


def main():
    parser = argparse.ArgumentParser(description="Microsoft docx file metadata extractor.")

    parser.add_argument("-m", "--media", dest="media", action="store_true", default=False,
                        help="Uncompress the stored media in the specified directory. " +
                             "As default, xtractor will create a directory with the name of the file.")

    parser.add_argument("-d", "--directory", dest="directory", action="store", default=".",
                        help="Name of the  directory where to output the media.")

    parser.add_argument(dest="item", action="store", metavar="[file or directory]",
                        help="DOCX file or directory to parse.")

    args = parser.parse_args()

    OfficeMetadata(args.item, args.directory, args.media)


if __name__ == '__main__':
    main()
