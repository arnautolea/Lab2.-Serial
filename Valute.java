package mvn.Lab2;

import com.thoughtworks.xstream.annotations.XStreamAlias;
import com.thoughtworks.xstream.annotations.XStreamAsAttribute;

@XStreamAlias("Valute")
public class Valute {

    @XStreamAlias("NumCode")
    private String numCode;
    @XStreamAlias("CharCode")
    private String сharCode;
    @XStreamAlias("Nominal")
    private int nominal;
    @XStreamAlias("Name")
    private String name;
    @XStreamAlias("Value")
    private double value;

    @XStreamAlias("ID")
    @XStreamAsAttribute
    private String id;

    @Override
    public String toString() {
        return "Valute{" +
                "numCode='" + numCode + '\'' +
                ", сharCode='" + сharCode + '\'' +
                ", nominal=" + nominal +
                ", name='" + name + '\'' +
                ", value=" + value +
                ", id='" + id + '\'' +
                '}';
    }

    public String getNumCode() {
        return numCode;
    }

    public void setNumCode(String numCode) {
        this.numCode = numCode;
    }

    public String getСharCode() {
        return сharCode;
    }

    public void setСharCode(String сharCode) {
        this.сharCode = сharCode;
    }

    public int getNominal() {
        return nominal;
    }

    public void setNominal(int nominal) {
        this.nominal = nominal;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public double getValue() {
        return value;
    }

    public void setValue(double value) {
        this.value = value;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }
}