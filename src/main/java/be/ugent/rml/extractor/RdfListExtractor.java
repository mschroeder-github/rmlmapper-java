
package be.ugent.rml.extractor;

import be.ugent.rml.NAMESPACES;
import be.ugent.rml.functions.SingleRecordFunctionExecutor;
import be.ugent.rml.records.Record;
import be.ugent.rml.store.Quad;
import be.ugent.rml.store.QuadStore;
import be.ugent.rml.term.NamedNode;
import be.ugent.rml.term.Term;
import java.util.ArrayList;
import java.util.List;

/**
 * Extracts from an rdf:List all the list items.
 */
public class RdfListExtractor implements Extractor, SingleRecordFunctionExecutor {

    private Term rdfList;
    private QuadStore store;

    public RdfListExtractor(Term rdfList, QuadStore store) {
        this.rdfList = rdfList;
        this.store = store;
    }
    
    @Override
    public List<Object> extract(Record record) {
        ArrayList<Object> result = new ArrayList<>();
        
        Term cur = rdfList;
        
        while(!cur.equals(new NamedNode(NAMESPACES.RDF + "nil"))) {
            List<Quad> firstQuads = store.getQuads(cur, new NamedNode(NAMESPACES.RDF + "first"), null);
            List<Quad> restQuads = store.getQuads(cur, new NamedNode(NAMESPACES.RDF + "rest"), null);
            
            if(!firstQuads.isEmpty()) {
                //just return a string because the NamedNodeGenerator will create NamedNodes from returned URIs
                result.add(firstQuads.get(0).getObject().getValue());
            }
            if(!restQuads.isEmpty()) {
                cur = restQuads.get(0).getObject();
            }
        }

        return result;
    }

    @Override
    public Object execute(Record record) throws Exception {
        return extract(record);
    }

}
