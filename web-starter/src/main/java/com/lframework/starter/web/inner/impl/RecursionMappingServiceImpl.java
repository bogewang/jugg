package com.lframework.starter.web.inner.impl;

import com.baomidou.mybatisplus.core.conditions.Wrapper;
import com.baomidou.mybatisplus.core.toolkit.Wrappers;
import com.google.common.collect.Lists;
import com.lframework.starter.common.constants.StringPool;
import com.lframework.starter.common.utils.CollectionUtil;
import com.lframework.starter.common.utils.StringUtil;
import com.lframework.starter.web.core.impl.BaseMpServiceImpl;
import com.lframework.starter.web.core.utils.ApplicationUtil;
import com.lframework.starter.web.core.utils.IdUtil;
import com.lframework.starter.web.inner.entity.RecursionMapping;
import com.lframework.starter.web.inner.enums.system.NodeType;
import com.lframework.starter.web.inner.mappers.RecursionMappingMapper;
import com.lframework.starter.web.inner.service.RecursionMappingService;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import lombok.NonNull;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

@Service
public class RecursionMappingServiceImpl extends BaseMpServiceImpl<RecursionMappingMapper, RecursionMapping>
        implements RecursionMappingService {

    @Override
    public List<String> getNodeParentIds(@NonNull String nodeId,
                                         @NonNull Class<? extends NodeType> nodeTypeClazz) {

        if (StringUtil.isEmpty(nodeId)) {
            return CollectionUtil.emptyList();
        }

        NodeType nodeType = ApplicationUtil.getBean(nodeTypeClazz);
        Wrapper<RecursionMapping> queryWrapper = Wrappers.lambdaQuery(RecursionMapping.class)
                .eq(RecursionMapping::getNodeId, nodeId)
                .eq(RecursionMapping::getNodeType, nodeType.getCode());

        RecursionMapping recursionMappings = getBaseMapper().selectOne(queryWrapper);
        if (recursionMappings == null || StringUtil.isEmpty(recursionMappings.getPath())) {
            return CollectionUtil.emptyList();
        }

        return StringUtil.split(recursionMappings.getPath(), StringPool.STR_SPLIT);
    }

    @Override
    public List<RecursionMapping> getParentNodes(List<String> nodeIdList, Class<? extends NodeType> nodeTypeClazz) {
        if (CollectionUtil.isEmpty(nodeIdList)) {
            return CollectionUtil.emptyList();
        }

        NodeType nodeType = ApplicationUtil.getBean(nodeTypeClazz);
        Wrapper<RecursionMapping> queryWrapper = Wrappers.lambdaQuery(RecursionMapping.class)
                .in(RecursionMapping::getNodeId, nodeIdList)
                .eq(RecursionMapping::getNodeType, nodeType.getCode());

        return getBaseMapper().selectList(queryWrapper);
    }

    @Override
    public List<RecursionMapping> selectAll(Class<? extends NodeType> nodeTypeClazz) {
        return getBaseMapper().selectList(Wrappers.lambdaQuery(RecursionMapping.class)
                .eq(RecursionMapping::getNodeType, nodeTypeClazz));
    }

    @Override
    public List<String> getNodeChildIds(String nodeId, Class<? extends NodeType> nodeTypeClazz) {

        NodeType nodeType = ApplicationUtil.getBean(nodeTypeClazz);
        return getBaseMapper().getNodeChildIds(nodeId, nodeType.getCode());
    }

    @Transactional(rollbackFor = Exception.class)
    @Override
    public void saveNode(String nodeId, Class<? extends NodeType> nodeTypeClazz) {

        this.saveNode(nodeId, nodeTypeClazz, null);
    }

    @Transactional(rollbackFor = Exception.class)
    @Override
    public void deleteNode(String nodeId, Class<? extends NodeType> nodeTypeClazz) {

        NodeType nodeType = ApplicationUtil.getBean(nodeTypeClazz);

        Wrapper<RecursionMapping> deleteWrapper = Wrappers.lambdaQuery(RecursionMapping.class)
                .eq(RecursionMapping::getNodeId, nodeId)
                .eq(RecursionMapping::getNodeType, nodeType.getCode());
        getBaseMapper().delete(deleteWrapper);
    }

    @Transactional(rollbackFor = Exception.class)
    @Override
    public void deleteBatchNode(List<String> nodeIdList, Class<? extends NodeType> nodeTypeClazz) {
        NodeType nodeType = ApplicationUtil.getBean(nodeTypeClazz);

        Wrapper<RecursionMapping> deleteWrapper = Wrappers.lambdaQuery(RecursionMapping.class)
                .in(RecursionMapping::getNodeId, nodeIdList)
                .eq(RecursionMapping::getNodeType, nodeType.getCode());
        getBaseMapper().delete(deleteWrapper);
    }

    @Transactional(rollbackFor = Exception.class)
    @Override
    public void deleteNodeAndChildren(String nodeId, Class<? extends NodeType> nodeTypeClazz) {
        NodeType nodeType = ApplicationUtil.getBean(nodeTypeClazz);
        List<String> childNodeIds = this.getNodeChildIds(nodeId, nodeTypeClazz);
        List<String> allNodeIds = new ArrayList<>();
        allNodeIds.add(nodeId);
        allNodeIds.addAll(childNodeIds);
        Wrapper<RecursionMapping> deleteWrapper = Wrappers.lambdaQuery(RecursionMapping.class)
                .in(RecursionMapping::getNodeId, allNodeIds)
                .eq(RecursionMapping::getNodeType, nodeType.getCode());
        getBaseMapper().delete(deleteWrapper);
    }

    @Transactional(rollbackFor = Exception.class)
    @Override
    public void saveNode(@NonNull String nodeId, @NonNull Class<? extends NodeType> nodeTypeClazz, List<String> parentIds) {

        NodeType nodeType = ApplicationUtil.getBean(nodeTypeClazz);
        Wrapper<RecursionMapping> queryWrapper = Wrappers.lambdaQuery(RecursionMapping.class)
                .eq(RecursionMapping::getNodeId, nodeId)
                .eq(RecursionMapping::getNodeType, nodeType.getCode());
        RecursionMapping data = getBaseMapper().selectOne(queryWrapper);
        if (data == null) {
            data = new RecursionMapping();
            data.setId(IdUtil.getId());
            data.setNodeId(nodeId);
            data.setNodeType(nodeType.getCode());
        }

        data.setLevel(1);
        data.setPath(StringPool.EMPTY_STR);
        if (!CollectionUtil.isEmpty(parentIds)) {
            data.setPath(StringUtil.join(StringPool.STR_SPLIT, parentIds));
            data.setLevel(parentIds.size() + 1);
        }

        getBaseMapper().insert(data);
    }

    @Override
    public void batchSaveNode(List<RecursionMapping> nodeList, Class<? extends NodeType> nodeTypeClazz) {
        if (CollectionUtil.isEmpty(nodeList)) {
            return;
        }
        List<String> NodeIdList = nodeList.stream().map(RecursionMapping::getNodeId).collect(Collectors.toList());

        NodeType nodeType = ApplicationUtil.getBean(nodeTypeClazz);
        Wrapper<RecursionMapping> queryWrapper = Wrappers.lambdaQuery(RecursionMapping.class)
                .in(RecursionMapping::getNodeId, NodeIdList)
                .eq(RecursionMapping::getNodeType, nodeType.getCode());
        // 获取所有的关联关系
        List<RecursionMapping> dbList = getBaseMapper().selectList(queryWrapper);
        Map<String, RecursionMapping> map = dbList.stream().collect(Collectors.toMap(RecursionMapping::getNodeId, r -> r));

        List<RecursionMapping> list = new ArrayList<>(nodeList.size());
        nodeList.forEach(item -> {
            RecursionMapping data = map.get(item.getNodeId());
            if (data == null) {
                data = new RecursionMapping();
                data.setId(IdUtil.getId());
                data.setNodeId(item.getNodeId());
                data.setNodeType(nodeType.getCode());
            }

            data.setLevel(1);
            data.setPath(StringPool.EMPTY_STR);
            if (StringUtil.isNotEmpty(item.getPath())) {
                data.setPath(item.getPath());
                List<String> strings = Lists.newArrayList(item.getPath().split(StringPool.STR_SPLIT));
                data.setLevel(strings.size() + 1);
            }
            list.add(data);
        });


        saveOrUpdateBatch(list);
    }
}
