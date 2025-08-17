from pydantic import BaseModel

class ItemBase(BaseModel):
    name: str
    description: str

class ItemCreate(ItemBase):
    pass

class ItemResponse(ItemBase):
    id: int

    class Config:
        from_attributes = True

class GroupDeleteRequest(BaseModel):
    item_ids: list[int]
    
    class Config:
        json_schema_extra = {
            "example": {
                "item_ids": [1, 2, 3]
            }
        }
    
    @classmethod
    def validate_item_ids(cls, v):
        if not v:
            raise ValueError("item_ids cannot be empty")
        if any(id <= 0 for id in v):
            raise ValueError("All item IDs must be positive integers")
        return v
