export function getCurrentContext(ctx) {
    return async dispatch => 
    {        
       dispatch({type:'CURRENTCONTEXT',data:ctx});        
    };
}
export function spContext(contextState={},action)
{
    switch (action.type) {
        case "CURRENTCONTEXT":            
            return action.data;    
        default:
        return contextState;
    }    
}