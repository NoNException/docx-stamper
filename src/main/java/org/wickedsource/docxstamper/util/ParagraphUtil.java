package org.wickedsource.docxstamper.util;

import org.docx4j.TraversalUtil;
import org.docx4j.finders.ClassFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.CTTxbxContent;
import org.docx4j.wml.CommentRangeEnd;
import org.docx4j.wml.CommentRangeStart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

public class ParagraphUtil {

    private static final ObjectFactory objectFactory = Context.getWmlObjectFactory();

    private ParagraphUtil() {

    }

    /**
     * Creates a new paragraph.
     *
     * @param texts the text of this paragraph. If more than one text is specified each text will be placed within its own Run.
     * @return a new paragraph containing the given text.
     */
    public static P create(String... texts) {
        P p = objectFactory.createP();
        for (String text : texts) {
            R r = RunUtil.create(text, p);
            p.getContent().add(r);
        }
        return p;
    }

    /**
     * Finds all Paragraphs in a Document which are in a TextBox
     * @param document
     * @return 
     */
    public static List<Object> getAllTextBoxes(WordprocessingMLPackage document) {
        ClassFinder finder = new ClassFinder(P.class); // docx4j class
        //necessary even if not used
        new TraversalUtil(document.getMainDocumentPart().getContent(),finder); // docx4j class
        ArrayList<Object> result = new ArrayList<>(finder.results.size());
        for (Object o : finder.results) {
            if (o instanceof P && ((P) o).getParent() instanceof CTTxbxContent) {
                result.add(o);
            }
        }
        return result;
    }

    public static List<P> getParagraphsInsideComment(P paragraph) {
        BigInteger commentId = null;
        boolean foundEnd = false;

        List<P> paragraphs = new ArrayList<>();
        paragraphs.add(paragraph);

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
                        if (next instanceof P) {
                            paragraphs.add((P) next);
                        }
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
        return paragraphs;
    }
}
