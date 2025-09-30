package org.SchedulesExtractor;


/**
 * Data from:
 * <a href="https://www.loc.gov/catdir/cpso/lcco/">
 *     Library of Congress Classification Outline
 * </a>
 */
public enum LOCSubject {
    A("General Works"),
    B("Philosophy. Psychology. Religion"),
    C("Auxiliary Sciences of History"),
    D("World History and History of Europe, Asia, Africa, Australia, New Zealand, Etc"),
    E("History of the Americas"),
    F("History of the Americas"),
    G("Geography. Anthropology. Recreation"),
    H("Social Sciences"),
    J("Political Science"),
    K("Law"),
    L("Education"),
    M("Music and Books On Music"),
    N("Fine Arts"),
    P("Language and Literature"),
    Q("Science"),
    R("Medicine"),
    S("Agriculture"),
    T("Technology"),
    U("Military Science"),
    V("Naval Science"),
    Z("Bibliography. Library Science. Information Resources (General)");

    private final String description;

    LOCSubject(String description) {
        this.description = description;
    }

    public String getDescription() {
        return description;
    }

    public static LOCSubject getLOCSubjectFromChar(char c) {
        char upperChar = Character.toUpperCase(c);

        for (LOCSubject subject : LOCSubject.values()) {
            if (subject.name().charAt(0) == upperChar) {
                return subject;
            }
        }
        return null;
    }
}