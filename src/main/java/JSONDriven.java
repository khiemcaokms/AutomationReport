import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

public class JSONDriven {



    public String getFeatureName(JsonArray array, int index){
        JsonObject object = array.get(index).getAsJsonObject();
        return object.get("name").toString().replaceAll("[^A-Za-z0-9]","");
    }

    public int getTotalScenario(JsonArray array, int index){
        JsonObject object = array.get(index).getAsJsonObject();
        JsonArray arrayScenarios = object.getAsJsonArray("elements");
        return arrayScenarios.size();
    }
    public String getScenarioName(JsonArray array, int feature, int scenario){
        JsonObject object = array.get(feature).getAsJsonObject();
        JsonArray arrayScenarios = object.getAsJsonArray("elements");
        JsonObject objectScenarios = arrayScenarios.get(scenario).getAsJsonObject();
        return objectScenarios.get("name").toString().replaceAll("[^A-Za-z0-9]","");
    }
    public String getScenarioStatus(JsonArray array, int feature, int scenario){
        JsonObject object = array.get(feature).getAsJsonObject();
        JsonArray arrayScenarios = object.getAsJsonArray("elements");
        JsonObject objectScenarios = arrayScenarios.get(scenario).getAsJsonObject();
        JsonArray arraySteps = objectScenarios.getAsJsonArray("steps");
        for (int i=0; i < arraySteps.size(); i++){
            JsonObject result = arraySteps.get(i).getAsJsonObject().get("result").getAsJsonObject();
            String status = result.get("status").toString();
            if (status.equalsIgnoreCase("\"passed\"")){
                continue;
            }else if(status.equalsIgnoreCase("\"failed\"")){
                return "Failed";
            }else if(status.equalsIgnoreCase("\"skipped\"")){
                return "Failed";
            }
        }
        return "Passed on Sauce Lab";
    }


}
