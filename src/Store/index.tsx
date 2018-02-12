import { Store,createStore, combineReducers, applyMiddleware } from 'redux';
import { reducer as reduxFormReducer } from 'redux-form';
import thunkMiddleware from 'redux-thunk';
import { createLogger } from 'redux-logger';
import { composeWithDevTools } from 'redux-devtools-extension';
import {spContext} from './../reducers/context';

const loggerMiddleware = createLogger();  //Remember to remove logger for production 
const reducer = combineReducers({
  form: reduxFormReducer, // mounted under "form" for redux-form default, currently not used
  spContext  : spContext 
});
const store = createStore(reducer,composeWithDevTools(
  applyMiddleware(thunkMiddleware,loggerMiddleware),
));
export default store;