package model;

/** Un punto della simulazione: percentuale + valori calcolati. */
public class SimulationPoint {

    private final double percent;        // es. -20, -19, ..., +20
    private final double posNoFix;        // POS senza compensazione
    private final double posTarget;       // POS target (costante)
    private final Double compensationVal; // valore della variabile compensata (qty o price)

    public SimulationPoint(double percent, double posNoFix, double posTarget, Double compensationVal) {
        this.percent = percent;
        this.posNoFix = posNoFix;
        this.posTarget = posTarget;
        this.compensationVal = compensationVal;
    }

    public double getPercent() { return percent; }
    public double getPosNoFix() { return posNoFix; }
    public double getPosTarget() { return posTarget; }
    public Double getCompensationVal() { return compensationVal; }
}
