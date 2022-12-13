package org.wickedsource.docxstamper.util;

import javax.xml.bind.JAXBElement;
import org.docx4j.finders.ClassFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.wml.CommentRangeEnd;
import org.docx4j.wml.CommentRangeStart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.wickedsource.docxstamper.api.DocxStamperException;
import org.wickedsource.docxstamper.replace.typeresolver.image.ImageResolver;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static java.util.function.Function.identity;
import static java.util.stream.Collectors.toList;

public class DocumentUtil {
    private DocumentUtil() {
        throw new DocxStamperException("Utility clases shouldn't be instantiated");
    }

    /**
     * Recursively walk through the content accessor to replace embedded images and import the matching
     * files to the destination document before importing content.
     *
     * @param sourceDocument document to import.
     * @param destDocument   document to add the source document content to.
     * @return the whole content of the source document with imported images replaced.
     * @throws Exception
     */
    public static List<Object> prepareDocumentForInsert(
            WordprocessingMLPackage sourceDocument,
            WordprocessingMLPackage destDocument
    ) throws Exception {
        return walkObjects(
                sourceDocument.getMainDocumentPart(),
                sourceDocument,
                destDocument);
    }

    /**
     * Recursively walk through the content accessor to replace embedded images and import the matching
     * files to the destination document.
     *
     * @param sourceContainer source container to walk.
     * @param sourceDocument  source document containing image files.
     * @param destDocument    destination document to add image files to.
     * @return the list of imported objects from the source container.
     * @throws Exception
     */
    private static List<Object> walkObjects(
            ContentAccessor sourceContainer,
            WordprocessingMLPackage sourceDocument,
            WordprocessingMLPackage destDocument
    ) throws Exception {
        List<Object> result = new ArrayList<>();
        for (Object obj : sourceContainer.getContent()) {
            if (obj instanceof R && isImageRun((R) obj)) {
                DocxImageExtractor docxImageExtractor = new DocxImageExtractor(sourceDocument);
                byte[] imageData = docxImageExtractor.getRunDrawingData((R) obj);
                String filename = docxImageExtractor.getRunDrawingFilename((R) obj);
                String alt = docxImageExtractor.getRunDrawingAltText((R) obj);
                Integer maxWidth = docxImageExtractor.getRunDrawingMaxWidth((R) obj);
                result.add(ImageResolver.createRunWithImage(destDocument, imageData, filename, alt, maxWidth));
            } else if (obj instanceof ContentAccessor) {
                List<Object> importedChildren = walkObjects((ContentAccessor) obj, sourceDocument, destDocument);
                ((ContentAccessor) obj).getContent().clear();
                ((ContentAccessor) obj).getContent().addAll(importedChildren);
                result.add(obj);
            } else {
                result.add(obj);
            }
        }
        return result;
    }


    /**
     * Check if a run contains an embedded image.
     *
     * @param run
     * @return true if the run contains an image, false otherwise.
     */
    private static boolean isImageRun(R run) {
        return run.getContent()
                .stream()
                .filter(runElement -> runElement instanceof JAXBElement)
                .map(JAXBElement.class::cast)
                .map(JAXBElement::getValue)
                .anyMatch(runValue -> runValue instanceof Drawing);
    }


    public static <T> List<T> extractElements(Object object, Class<T> elementClass) {
        // we handle full documents slightly differently as they have headers and footers
        if (object instanceof WordprocessingMLPackage) {
            WordprocessingMLPackage document = (WordprocessingMLPackage) object;

            return Stream.of(
                            getElementStreamFrom(document, elementClass, Namespaces.HEADER),
                            getElementStream(document.getMainDocumentPart().getJaxbElement(), elementClass),
                            getElementStreamFrom(document, elementClass, Namespaces.FOOTER)
                    )
                    .flatMap(identity())
                    .collect(toList());
        }

        return getElementStream(object, elementClass)
                .collect(toList());
    }

    private static <T> Stream<T> getElementStreamFrom(
            WordprocessingMLPackage document,
            Class<T> clazz,
            String relationshipType
    ) {
        RelationshipsPart relationshipsPart = document
                .getMainDocumentPart()
                .getRelationshipsPart();
        return relationshipsPart
                .getRelationships()
                .getRelationship()
                .stream()
                .filter(relationship -> relationship.getType().equals(relationshipType))
                .map(relationshipsPart::getPart)
                .flatMap(relationshipPart -> getElementStream(relationshipPart, clazz));
    }

    private static <T> Stream<T> getElementStream(Object obj, Class<T> clazz) {
        ClassFinder finder = new ClassFinder(clazz);
        finder.walkJAXBElements(obj);
        return finder.results
                .stream()
                .map(clazz::cast);
    }

    public static List<Object> getDocPartsInsideComment(P paragraph) {
        BigInteger commentId = null;
        boolean foundEnd = false;

        List<Object> docParts = new ArrayList<>();
        docParts.add(paragraph);

        for (Object object : paragraph.getContent()) {
            if (object instanceof CommentRangeStart) {
                commentId = ((CommentRangeStart) object).getId();
            }
            if (object instanceof CommentRangeEnd && commentId != null && commentId.equals(((CommentRangeEnd) object).getId())) {
                foundEnd = true;
            }
        }
        if (!foundEnd && commentId != null) {
            Object parent = paragraph.getParent();
            if (parent instanceof ContentAccessor) {
                ContentAccessor contentAccessor = (ContentAccessor) parent;
                int index = contentAccessor.getContent().indexOf(paragraph);
                for (int i = index + 1; i < contentAccessor.getContent().size() && !foundEnd; i++) {
                    Object next = contentAccessor.getContent().get(i);

                    if (next instanceof CommentRangeEnd && ((CommentRangeEnd) next).getId().equals(commentId)) {
                        foundEnd = true;
                    } else {
                        docParts.add(next);
                        if (next instanceof ContentAccessor) {
                            ContentAccessor childContent = (ContentAccessor) next;
                            for (Object child : childContent.getContent()) {
                                if (child instanceof CommentRangeEnd && ((CommentRangeEnd) child).getId().equals(commentId)) {
                                    foundEnd = true;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
        return docParts;
    }

}
