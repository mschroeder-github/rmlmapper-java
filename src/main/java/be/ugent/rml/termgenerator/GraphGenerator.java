
package be.ugent.rml.termgenerator;

import be.ugent.rml.NAMESPACES;
import be.ugent.rml.functions.FunctionUtils;
import be.ugent.rml.functions.SingleRecordFunctionExecutor;
import be.ugent.rml.records.Record;
import be.ugent.rml.store.Quad;
import be.ugent.rml.store.QuadStore;
import be.ugent.rml.store.RDF4JStore;
import be.ugent.rml.term.NamedNode;
import be.ugent.rml.term.Term;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import org.apache.commons.io.IOUtils;
import org.eclipse.rdf4j.rio.RDFFormat;

/**
 * A function can return turtle syntax.
 * The returned triples are added to the executor's resultingQuads and selected
 * resource of the graph are returned as terms. This way one record can be mapped to
 * a complex RDF graph. This helps when the record contains more than one
 * thing, for exmaple when open knowledge extraction is applied to a sentence in
 * a CSV cell.
 */
public class GraphGenerator extends TermGenerator {

    //because a graph was generated we put the triples directly 
    //in the resulting quads store
    private QuadStore resultingQuads;
    
    public GraphGenerator(SingleRecordFunctionExecutor functionExecutor, QuadStore resultingQuads) {
        super(functionExecutor);
        this.resultingQuads = resultingQuads;
    }

    @Override
    public List<Term> generate(Record record) throws Exception {
        List<String> objectStrings = new ArrayList<>();
        FunctionUtils.functionObjectToList(functionExecutor.execute(record), objectStrings);
        
        ArrayList<Term> objects = new ArrayList<>();

        for(String objStr : objectStrings) {
            
            //read turtle
            QuadStore qs = new RDF4JStore();
            qs.read(IOUtils.toInputStream(objStr, StandardCharsets.UTF_8), null, RDFFormat.TURTLE);
            
            //a special resource (may be BlankNode) of type ss:SelectedObjects that points to the returning terms
            for(Quad selectedObjects : qs.getQuads(null, new NamedNode(NAMESPACES.RDF + "type"), new NamedNode(NAMESPACES.SS + "SelectedObjects"))) {
                //we reuse rr:object here
                for(Quad termQuad : qs.getQuads(selectedObjects.getSubject(), new NamedNode(NAMESPACES.RR + "object"), null)) {
                    
                    //only the referred objects will be returned
                    objects.add(termQuad.getObject());
                    
                    //remove, so it can not be in the resultingQuads graph
                    qs.removeQuads(termQuad);
                }
                
                //remove, so it can not be in the resultingQuads graph
                qs.removeQuads(selectedObjects);
            }
            
            //the rest is directly added to the resulting quads graph
            for(Quad q : qs.getQuads(null, null, null)) {
                resultingQuads.addQuad(q);
            }
        }
        
        return objects;
    }

}
