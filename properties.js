export default class Properties {
    static get(key) {
      // dummy for node
      if (typeof PropertiesService === 'undefined') {
        return 'test';
      }
      return PropertiesService.getScriptProperties().getProperty(key);
    }
  
    static set(key, value) {
      // dummy for node
      if (typeof PropertiesService === 'undefined') {
        return;
      }
      PropertiesService.getScriptProperties().setProperty(key, value);
    }
  
    static getAll() {
      // dummy for node
      if (typeof PropertiesService === 'undefined') {
        return {};
      }
      return PropertiesService.getScriptProperties().getProperties();
    }
  }
  