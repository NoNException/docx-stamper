package org.wickedsource.docxstamper.processor.repeat;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.CommentRangeEnd;
import org.docx4j.wml.CommentRangeStart;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.P;
import org.wickedsource.docxstamper.DocxStamperConfiguration;
import org.wickedsource.docxstamper.api.typeresolver.TypeResolverRegistry;
import org.wickedsource.docxstamper.el.ExpressionResolver;
import org.wickedsource.docxstamper.processor.BaseCommentProcessor;
import org.wickedsource.docxstamper.replace.PlaceholderReplacer;
import org.wickedsource.docxstamper.util.CommentUtil;
import org.wickedsource.docxstamper.util.ParagraphUtil;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.wickedsource.docxstamper.util.ParagraphUtil.getParagraphsInsideComment;

public class ParagraphRepeatProcessor extends BaseCommentProcessor implements IParagraphRepeatProcessor {

    private static class ParagraphsToRepeat {
        List<Object> data;
        List<P> paragraphs;
    }

    private Map<P, ParagraphsToRepeat> pToRepeat = new HashMap<>();

    private final PlaceholderReplacer<Object> placeholderReplacer;
    private final DocxStamperConfiguration config;

    public ParagraphRepeatProcessor(TypeResolverRegistry typeResolverRegistry, ExpressionResolver expressionResolver, DocxStamperConfiguration config) {
        this.placeholderReplacer = new PlaceholderReplacer<>(typeResolverRegistry);
        this.config = config;
        this.placeholderReplacer.setExpressionResolver(expressionResolver);
    }

    @Override
    public void repeatParagraph(List<Object> objects) {

        P paragraph = getParagraph();
        List<P> paragraphs = getParagraphsInsideComment(paragraph);

        ParagraphsToRepeat toRepeat = new ParagraphsToRepeat();
        toRepeat.data = objects;
        toRepeat.paragraphs = paragraphs;

        pToRepeat.put(paragraph, toRepeat);
        CommentUtil.deleteComment(getCurrentCommentWrapper());
    }

    @Override
    public void commitChanges(WordprocessingMLPackage document) {
        for (Map.Entry<P, ParagraphsToRepeat> entry : pToRepeat.entrySet()) {
            P paragraph = entry.getKey();
            ParagraphsToRepeat paragraphsToRepeat = entry.getValue();
            List<Object> expressionContexts = paragraphsToRepeat.data;

            List<P> paragraphsToAdd = new ArrayList<>();

            if (expressionContexts == null) {
                if (config.isReplaceNullValues() && config.getNullValuesDefault() != null) {
                    P nullReplacedParagraph = ParagraphUtil.create(config.getNullValuesDefault());
                    paragraphsToAdd.add(nullReplacedParagraph);
                }
            } else for (Object expressionContext : expressionContexts) {
                for (P paragraphToClone : paragraphsToRepeat.paragraphs) {
                    P pClone = XmlUtils.deepCopy(paragraphToClone);
                    placeholderReplacer.resolveExpressionsForParagraph(pClone, expressionContext, document);
                    paragraphsToAdd.add(pClone);
                }
            }

            Object parent = paragraph.getParent();
            if (parent instanceof ContentAccessor) {
                ContentAccessor contentAccessor = (ContentAccessor) parent;
                int index = contentAccessor.getContent().indexOf(paragraph);
                if (index >= 0) {
                    contentAccessor.getContent().addAll(index, paragraphsToAdd);
                }
                contentAccessor.getContent().removeAll(paragraphsToRepeat.paragraphs);
            }
        }
    }

    @Override
    public void reset() {
        pToRepeat = new HashMap<>();
    }

}
