-------------Version_1----------------
1. base version to generate ppt from the LLM response.
2. using cortext complet() stateless inference call to convert the long text replies from the cortext agent to bullet points for PPT representation purposes.
3. Future versions will try not to use the llm inference call for generating bullet points.
4. For now the Charts created in the PPT are created extracting labels + values and rebuilding a generic PowerPoint chart - (x/y appear interchanged compared to how it is showed in the Frontend as they are image rendered with vega)
5. Will add PPT creation suppport for all types of scenrios in the next versions - (In this version the charts created in the PPT are sometimes missing some of the Items if the Num of items is too long )
6. Will add support for accepting multipe messages for PPT generation with multiple individual slides when sent as a collection from front end. 
7. Will replace geenric parser with the vega lite compatable parser for better chart generation in PPT.
