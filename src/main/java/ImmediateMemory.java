public class ImmediateMemory {
    private byte listLearningScore;
    private byte storyMemoryScore;

    public ImmediateMemory(byte listLearningScore, byte storyMemoryScore) {
        this.listLearningScore = listLearningScore;
        this.storyMemoryScore = storyMemoryScore;
    }

    public byte getListLearningScore() {
        return listLearningScore;
    }

    public void setListLearningScore(byte listLearningScore) {
        this.listLearningScore = listLearningScore;
    }

    public byte getStoryMemoryScore() {
        return storyMemoryScore;
    }

    public void setStoryMemoryScore(byte storyMemoryScore) {
        this.storyMemoryScore = storyMemoryScore;
    }
}
