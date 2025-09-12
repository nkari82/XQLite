-- 샘플 데이터
INSERT INTO items (id,row_version,deleted,name,rarity,required_lv,atk_lv50,options)
VALUES
(1,1,0,'Bronze Sword',1,1,50,'{"tags":["sword","starter"],"vendor":"village"}'),
(2,2,0,'Iron Sword',2,5,120,'{"tags":["sword","iron"],"vendor":"blacksmith"}'),
(3,3,0,'Flame Dagger',3,8,110,'{"tags":["dagger","fire"],"vendor":"dungeon"}'),
(4,4,0,'Knight Blade',4,12,180,'{"tags":["sword","elite"],"vendor":"castle"}'),
(5,5,0,'Dragon Slayer',5,18,260,'{"tags":["sword","legend"],"vendor":"raid"}');
