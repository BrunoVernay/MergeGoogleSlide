package link.brnvrn.com;

import java.util.*;

/**
 * @TODO find a better name
 */
public class Linko {
    private String field;       // Text in the template slide (ex: "{{name}}")
    private String objctId;     // ObjectId that the program will find in the slide (you don't need to know it)
    private int column;         // The column index in the spreadsheet

    Set<Linko> link = new LinkedHashSet<>(10);

    /**
     *
     * @param field
     * @param column
     */
    private Linko(String field, int column) {
        this.field = field;
        this.column = column;
    }

    public Linko() {
    }

    /**
     * Used to declare the link between the column and the text in template slide
     * @param field
     * @param column
     */
    public void add(String field, int column) {
        link.add(new Linko(field, column));
    }

    /**
     * Set the initials ObjectsId where we find the field text.
     * @param content
     * @param objectId
     * @return
     */
    public boolean setMatch(String content, String objectId) {
        for (Linko l: link) {
            if (content.contains(l.field)) {
                l.objctId = objectId;
                return true;
            }
        }
        return false;
    }

    public String[] getFields() {
        List<String> result = new ArrayList<>();
        for(Linko l: link) {
            result.add(l.field);
        }
        return result.toArray(new String[0]);
    }

    /**
     * get the mapping to duplicate the slides
     * @param i row loop index starting at 0
     * @return
     */
    public Map<String,String> getMapId(int i) {
        Map<String,String> result = new LinkedHashMap<>();
        for(Linko l: link) {
            String oldId;
            String newId = l.objctId+"-BVE-"+i;
            if (i==0)
                 oldId = l.objctId;
            else oldId = l.objctId+"-BVE-"+(i-1);
            result.put(oldId, newId);
        }
        return result;
    }

    /**
     * Get the mapping of "old" ObjectId and the column containing the text to insert
     * @param i row loop index starting at 0
     * @return
     */
    public Iterable<? extends Map.Entry<String, Integer>> getMapObj2Col(int i) {
        Map<String,Integer> result = new LinkedHashMap<>();
        for(Linko l: link) {
            String oldId;
            if (i==0)
                oldId = l.objctId;
            else oldId = l.objctId+"-BVE-"+(i-1);
            result.put(oldId, l.column);
        }
        return result.entrySet();
    }
}
