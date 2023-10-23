public class Subject {
    private String id;
    private byte group;
    private byte age;
    private TestingStage testingStage;

    private ImmediateMemory immediateMemory;
    private VisuospatialConstructional visConstr;
    private Language language;
    private Attention attention;
    private DelayedMemory delayedMemory;

    public Subject(String id, byte group, byte age, TestingStage testingStage) {
        this.id = id;
        this.group = group;
        this.age = age;
        this.testingStage = testingStage;
    }

    public String getId() {
        return id;
    }

    public byte getGroup() {
        return group;
    }

    public byte getAge() {
        return age;
    }

    public TestingStage getTestingStage() {
        return testingStage;
    }

    public ImmediateMemory getImmediateMemory() {
        return immediateMemory;
    }

    public VisuospatialConstructional getVisConstr() {
        return visConstr;
    }

    public Language getLanguage() {
        return language;
    }

    public Attention getAttention() {
        return attention;
    }

    public DelayedMemory getDelayedMemory() {
        return delayedMemory;
    }
}
